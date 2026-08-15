#!/usr/bin/env python3
"""Shipping Enrichment Lite - Phase 1 + Phase 2 + tiny live pilot + browser mode.

Reads skus.csv, produces an Excel workbook plus CSV review files, in a run folder.

For each unique ASIN, classification comes from (in priority order):
  1. rendered_cache/<ASIN>.html (from an earlier --browser-live run), reparsed -
     always preferred when present, since a rendered page is a superset of a raw one.
  2. page_cache/<ASIN>.html (from an earlier --live run), reparsed.
  3. --browser-live only, no cache: a single headless-Chrome DOM dump (real
     Amazon AU page, JS executed) - no proxy, no concurrency - then cached and
     parsed the same way as a fixture.
  4. --live only, no cache: a single, direct, uncached raw fetch - no proxy, no
     Decodo, no concurrency - then cached and parsed the same way.
  5. Otherwise: the offline fixtures/<ASIN>.html lookup, then the built-in mock
     table, then a clear "nothing available" marker.

Network access happens ONLY when --live or --browser-live is passed, and even
then at most once per ASIN per run, direct (no proxy), never concurrent, with a
mandatory sleep between requests and a hard cap via --limit/--browser-limit.
See README.md for usage.

Usage:
    python3 run_shipping_enrichment.py [--input skus.csv] [--fixtures fixtures] [--mock-only]
    python3 run_shipping_enrichment.py --live --limit 5 --delay-seconds 5
    python3 run_shipping_enrichment.py --browser-live --browser-limit 1 --browser-delay-seconds 5
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import gzip
import html
import http.cookiejar
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zlib
from collections import OrderedDict
from pathlib import Path

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
RESULTS_DIR = os.path.join(PROJECT_DIR, "results")
RUNS_DIR = os.path.join(PROJECT_DIR, "runs")
FIXTURES_DIR = os.path.join(PROJECT_DIR, "fixtures")
CACHE_DIR = os.path.join(PROJECT_DIR, "page_cache")
RENDERED_CACHE_DIR = os.path.join(PROJECT_DIR, "rendered_cache")

DELIVERY_LOCATION_BASIS = "AU postcode 2000, logged-out, non-Prime"
ASIN_RE = re.compile(r"^[A-Z0-9]{10}$")

# The complete output contract. Deliberately minimal: SKU/ASIN -> ETA, merchant,
# delivery fee, plus provenance. Nothing is collected "for later" - no title,
# brand, images, ratings, reviews, categories, specs, variations, or seller
# profile data is ever parsed or stored.
COLUMNS = [
    "sku",
    "asin",
    "eta_raw",
    "eta_date",
    "merchant",
    "delivery_fee",
    "checked_at",
    "status",
    "error",
    # Not "extra data" - this is what makes the three fields above interpretable.
    # A logged-out session is geolocated by exit IP, so an ETA/fee can silently be
    # for the wrong country. Shipping an ETA without saying where it was measured
    # from is how a negative-profit postage policy gets written. See CLAUDE.md.
    "delivery_location",
]

# status -> sheet it belongs on. See CLAUDE.md: nothing is silently dropped.
SHEET_FOR_STATUS = {
    "extracted": "extracted",
    "redirect": "redirects",
    "no_delivery_info": "no_required_data",
    "no_delivery_fee": "no_required_data",
    "no_merchant": "no_required_data",
    "out_of_stock": "no_required_data",
    "invalid_row": "errors",
    "not_found": "errors",
    "server_busy": "errors",
    "proxy_auth_failed": "errors",
    "timeout": "errors",
    "unsupported_layout": "errors",
    "fetch_failed": "errors",
    "parse_error": "errors",
}
DATA_SHEETS = ["extracted", "redirects", "no_required_data", "errors"]


# --------------------------------------------------------------------------
# Input loading
# --------------------------------------------------------------------------

def normalise_asin(raw):
    """Uppercase, trim, strip Excel's leading-apostrophe artefact."""
    return (raw or "").strip().lstrip("'").upper()


def load_input(path, log):
    """Read the input CSV.

    Returns (rows, stats). Every non-blank row gets an input_row_number assigned
    before any filtering, so a row can always be traced back to the spreadsheet.
    Invalid rows are kept with a rejection reason rather than dropped.
    """
    rows = []
    blank_skipped = 0

    with open(path, newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        if not reader.fieldnames:
            raise SystemExit("Input file has no header row: %s" % path)
        headers = {(h or "").strip().lower(): h for h in reader.fieldnames}
        for key in ("sku", "asin"):
            if key not in headers:
                raise SystemExit("Input file is missing required column: %s" % key)

        row_number = 0
        for raw in reader:
            values = [(v or "").strip() for v in raw.values()]
            if not any(values):
                blank_skipped += 1
                continue

            row_number += 1
            sku = (raw.get(headers["sku"]) or "").strip()
            asin = normalise_asin(raw.get(headers["asin"]))
            rc_listing_id = (raw.get(headers.get("rc_listing_id", "")) or "").strip()
            title = (raw.get(headers.get("title", "")) or "").strip()

            if not asin and sku and ASIN_RE.match(normalise_asin(sku)):
                # RC sometimes supplies a bare list of ASINs with no separate SKU
                # column (see PLAN.md section 16, Q1 - exact input format still
                # unconfirmed). A row with only one value in the "sku" position
                # that is itself ASIN-shaped is that case, not a missing ASIN -
                # treat it as the ASIN and echo it as the SKU too.
                asin = normalise_asin(sku)

            reason = ""
            if not sku and not asin:
                reason = "missing_sku_and_asin"
            elif not asin:
                reason = "missing_asin"
            elif not ASIN_RE.match(asin):
                reason = "invalid_asin"
            elif not sku:
                reason = "missing_sku"

            rows.append({
                "input_row_number": row_number,
                "sku": sku,
                "asin": asin,
                "rc_listing_id": rc_listing_id,
                "title": title,
                "reject_reason": reason,
            })

    log("read %d data rows (%d blank rows skipped as padding)" % (len(rows), blank_skipped))
    return rows, {"blank_rows_skipped": blank_skipped}


def dedupe(rows, log):
    """Group valid rows by fetch key: ASIN when present, otherwise SKU.

    One fetch unit per unique key. The result is fanned back out to every row that
    referenced it, so every input row still receives an answer.
    """
    units = OrderedDict()
    for row in rows:
        if row["reject_reason"]:
            continue
        key = row["asin"] or row["sku"]
        units.setdefault(key, []).append(row)

    valid = sum(len(v) for v in units.values())
    saved = valid - len(units)
    log("deduped %d valid rows into %d fetch units (%d duplicate fetches avoided)"
        % (valid, len(units), saved))
    return units, saved


# --------------------------------------------------------------------------
# Mock classification - PHASE 1 ONLY, no network
# --------------------------------------------------------------------------

def _blank_record():
    return {c: "" for c in COLUMNS}


MOCK_RESULTS = {
    # Two clean extractions, deliberately different so the ETA flag is exercised.
    "B0TEST0001": {
        "status": "extracted",
        "reason": "",
        "merchant": "Mock Seller Pty Ltd",
        "sold_by": "Mock Seller Pty Ltd",
        "shipped_by": "Amazon AU",
        "ships_from": "Amazon AU",
        "delivered_by": "Amazon AU",
        "delivery_message_raw": "FREE delivery Wednesday, 13 August",
        "delivery_fee": 0.0,
        "delivery_fee_currency": "AUD",
        "eta_date_text": "Wednesday, 13 August",
        "eta_days_max": 9,
        "eta_over_14_days": "FALSE",
        "availability": "IN_STOCK",
    },
    "B0TEST0002": {
        "status": "extracted",
        "reason": "",
        "merchant": "Mock Marketplace Trader",
        "sold_by": "Mock Marketplace Trader",
        "shipped_by": "Mock Marketplace Trader",
        "ships_from": "Overseas",
        "delivered_by": "Mock Marketplace Trader",
        "delivery_message_raw": "$9.99 delivery Monday, 25 August",
        "delivery_fee": 9.99,
        "delivery_fee_currency": "AUD",
        "eta_date_text": "Monday, 25 August",
        "eta_days_max": 21,
        "eta_over_14_days": "TRUE",
        "availability": "IN_STOCK",
    },
    # Canonicalises elsewhere. Recorded, not followed - see PLAN.md section 9.
    "B0TEST0003": {
        "status": "redirect",
        "reason": "canonical_asin_differs",
        "final_asin": "B0TESTCAN1",
    },
    # Fetched fine, but the fee could not be determined. Blank fee, NOT zero.
    "B0TEST0004": {
        "status": "no_delivery_fee",
        "reason": "delivery_block_present_no_fee_parsed",
        "merchant": "Mock Seller Pty Ltd",
        "sold_by": "Mock Seller Pty Ltd",
        "shipped_by": "Mock Seller Pty Ltd",
        "ships_from": "Sydney, NSW",
        "delivered_by": "Mock Seller Pty Ltd",
        "delivery_message_raw": "Delivery information unavailable for this offer",
        "availability": "IN_STOCK",
    },
    # Infrastructure fault. Says nothing about the product.
    "B0TEST0005": {
        "status": "server_busy",
        "reason": "amazon_throttle_stub_returned",
    },
}

DEFAULT_MOCK = {"status": "no_delivery_info", "reason": "fixture_missing_no_mock_fallback"}


# --------------------------------------------------------------------------
# Fixture parsing - PHASE 2, offline only, no network
#
# Conservative regex-based extraction over saved HTML. Not a full HTML parser: it
# assumes ids are unique and tags are reasonably well-formed, using a balanced-tag
# scan keyed off id="..." attributes. That's enough for the synthetic fixtures here
# and for a single saved Amazon page, but a genuinely malformed page can defeat it -
# which is exactly why every field here defaults to "unknown" (blank) rather than
# raising, and a top-level try/except in parse_fixture() routes anything unexpected
# to errors instead of crashing the run. See fixtures/README.md.
# --------------------------------------------------------------------------

_CANONICAL_ASIN_RE = re.compile(
    r'<link[^>]+rel=["\']canonical["\'][^>]*href=["\'][^"\']*?/dp/([A-Za-z0-9]{10})["\']',
    re.IGNORECASE,
)

_MONTHS = ("January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December")
_DAY_NAMES = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
_ETA_DATE_RE = re.compile(
    r'\b(%s),?\s+(\d{1,2})\s+(%s)\b' % ("|".join(_DAY_NAMES), "|".join(_MONTHS))
)
_FEE_RE = re.compile(r'\$\s?(\d+(?:\.\d{1,2})?)')


def extract_canonical_asin(html_text):
    """Requested vs final ASIN. See PLAN.md section 9: a redirect is recorded, not
    auto-followed, so this is as far as parsing goes when it fires."""
    match = _CANONICAL_ASIN_RE.search(html_text)
    return match.group(1).upper() if match else None


def _strip_tags(fragment):
    text = re.sub(r"<[^>]+>", " ", fragment)
    text = html.unescape(text)
    # Amazon pads some widget labels with zero-width joiners/spaces.
    text = re.sub(r"[​-‏﻿]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    # An inline tag (e.g. <b>date</b>.) leaves a stray space before punctuation
    # once its tags are blanked out - tidy that up, it's cosmetic only.
    return re.sub(r"\s+([,.;:!?])", r"\1", text)


def extract_element_text_by_id(html_text, element_id):
    """Stripped text content of the first element (any tag) with this id attribute.

    Balanced-tag scan, not a full parser: finds the opening tag, then counts
    same-tag open/close pairs until depth returns to zero. Returns None (never
    raises) if the id isn't found or the markup around it is unbalanced - both
    cases mean "unknown", not "error".
    """
    open_re = re.compile(r'<([a-zA-Z][a-zA-Z0-9]*)\b[^>]*\bid=["\']%s["\'][^>]*>'
                          % re.escape(element_id))
    open_match = open_re.search(html_text)
    if not open_match:
        return None
    tag = open_match.group(1)
    start = open_match.end()
    depth = 1
    tag_re = re.compile(r'<(/?)%s\b[^>]*>' % re.escape(tag), re.IGNORECASE)
    for tag_match in tag_re.finditer(html_text, start):
        if tag_match.group(1):  # closing tag
            depth -= 1
            if depth == 0:
                return _strip_tags(html_text[start:tag_match.start()])
        else:
            depth += 1
    return None


def _collapse_repeat(text):
    """Amazon renders the offer value twice (visible + truncation popover), so the
    naive text of the block reads 'Amazon AU Amazon AU'. Collapse an exact
    whole-string repetition back to one copy. Leaves genuinely distinct text alone.
    """
    words = text.split()
    n = len(words)
    for size in range(1, n // 2 + 1):
        if n % size:
            continue
        chunk = words[:size]
        if all(words[i:i + size] == chunk for i in range(0, n, size)):
            return " ".join(chunk)
    return text


# Real markup for one offer row reads: "<value> <value> <label> <value>" - the
# value is rendered visibly, again inside a truncation popover, and a third time
# after the label within that popover. Stripping the label ANYWHERE (not just at
# the start) leaves a clean n-times repetition that _collapse_repeat reduces to
# one copy. Verified against live cached pages for both Amazon-sold and
# third-party offers.
_OFFER_LABELS_RE = re.compile(
    r"\b(Shipper\s*/\s*Seller|Sold by|Ships from|Dispatched from)\b\s*[:\-]?\s*",
    re.IGNORECASE)


def _extract_offer_feature_value(html_text, feature_name):
    """Value of one buy-box offer row (e.g. desktop-merchant-info).

    Targets the 'offer-display-feature-text' div specifically, so the row's
    LABEL div ('Shipper / Seller') is never mistaken for the value - that bug
    produced merchant='Shipper / Seller Amazon AU Amazon AU ...' on real pages.
    """
    pattern = re.compile(
        r'<div[^>]*class="[^"]*offer-display-feature-text[^"]*"[^>]*'
        r'offer-display-feature-name="%s"[^>]*>' % re.escape(feature_name),
        re.IGNORECASE)
    match = pattern.search(html_text)
    if not match:
        return None
    start = match.end()
    depth = 1
    for tag_match in re.finditer(r'<(/?)div\b[^>]*>', html_text[start:]):
        depth += -1 if tag_match.group(1) else 1
        if depth == 0:
            text = _strip_tags(html_text[start:start + tag_match.start()])
            text = re.sub(r"\s+", " ", _OFFER_LABELS_RE.sub(" ", text)).strip()
            text = _collapse_repeat(text).strip()
            return text or None
    return None


def parse_merchant_fields(html_text):
    """Returns (sold_by, ships_from), either may be None. Never invents a value."""
    sold_by = _extract_offer_feature_value(html_text, "desktop-merchant-info")
    if not sold_by:
        sold_by = extract_element_text_by_id(html_text, "sellerProfileTriggerId")
    if not sold_by:
        block = extract_element_text_by_id(html_text, "merchantInfoFeature_feature_div")
        if block:
            block = re.sub(r"\s+", " ", _OFFER_LABELS_RE.sub(" ", block)).strip()
            sold_by = _collapse_repeat(block).strip() or None

    ships_from = _extract_offer_feature_value(html_text, "desktop-fulfiller-info")
    if not ships_from:
        block = extract_element_text_by_id(html_text, "fulfillerInfoFeature_feature_div")
        if block:
            block = re.sub(r"\s+", " ", _OFFER_LABELS_RE.sub(" ", block)).strip()
            ships_from = _collapse_repeat(block).strip() or None

    return sold_by, ships_from


def _days_until(day, month_name, today=None):
    """Day-count from today to the nearest future (or today) occurrence of this
    day/month. No year appears on an Amazon delivery message, so this assumes the
    soonest calendar match. Returns None only if day/month can't form a valid date
    (e.g. 31 February) - defensive, not expected to fire on real pages."""
    today = today or dt.date.today()
    month = _MONTHS.index(month_name) + 1
    for year in (today.year, today.year + 1):
        try:
            candidate = dt.date(year, month, day)
        except ValueError:
            continue
        if candidate >= today:
            return (candidate - today).days
    return None


def parse_delivery(html_text):
    """Returns a dict with whichever of delivery_message_raw / delivery_fee /
    delivery_fee_currency / eta_date_text / eta_days_max / eta_over_14_days could
    be determined unambiguously. Missing keys mean "leave that column blank" -
    the caller never fills them with a guessed default.
    """
    raw = extract_element_text_by_id(html_text, "deliveryBlockMessage")
    if raw is None:
        raw = extract_element_text_by_id(html_text, "mir-layout-DELIVERY_BLOCK")
    if not raw:
        return {}

    result = {"delivery_message_raw": raw}

    if re.search(r"\bfree\b", raw, re.IGNORECASE):
        # FREE delivery is a fee of 0, not "unknown". See CLAUDE.md: blank is not zero,
        # and the inverse matters just as much - zero must not become blank.
        result["delivery_fee"] = 0.0
        result["delivery_fee_currency"] = "AUD"
    else:
        fee_match = _FEE_RE.search(raw)
        if fee_match:
            result["delivery_fee"] = float(fee_match.group(1))
            result["delivery_fee_currency"] = "AUD"
        # else: no unambiguous fee in the text. Leave delivery_fee/currency unset,
        # which build_records() renders as blank - never invent a number.

    eta_date = normalise_eta_date(raw)
    if eta_date:
        result["eta_date"] = eta_date
    # else: leave eta_date unset -> blank. eta_raw still carries the original
    # wording, so nothing is lost. Never guess a date.

    return result


# Amazon AU delivery wording, in the shapes actually seen on real pages:
#   "FREE delivery Wednesday, 20 May"
#   "Delivery 24 August - 2 September"      (range -> take the LATEST date)
#   "Arrives 3 - 8 September"                (shared month on the right)
# Weekday is optional; day/month order varies. A range always resolves to its
# maximum, because the business question is "how late might this arrive?".
_MONTH_NAMES = "|".join(_MONTHS) + "|" + "|".join(m[:3] for m in _MONTHS)
_DATE_DM_RE = re.compile(r'\b(\d{1,2})\s+(%s)\b' % _MONTH_NAMES, re.IGNORECASE)
_DATE_MD_RE = re.compile(r'\b(%s)\s+(\d{1,2})\b' % _MONTH_NAMES, re.IGNORECASE)
# "3 - 8 September" / "3 to 8 September": leading bare day sharing a later month.
_BARE_RANGE_RE = re.compile(
    r'\b(\d{1,2})\s*(?:-|–|to)\s*(\d{1,2})\s+(%s)\b' % _MONTH_NAMES, re.IGNORECASE)


def _month_number(name):
    name = name.strip().lower()
    for index, month in enumerate(_MONTHS, start=1):
        if month.lower() == name or month[:3].lower() == name:
            return index
    return None


def _resolve_future_date(day, month, today=None):
    """Nearest occurrence of day/month that is not in the past. Amazon omits the
    year, so this assumes the soonest match - correct for delivery estimates,
    which are always near-future. Returns a date or None if invalid (e.g. 31 Feb)."""
    today = today or dt.date.today()
    for year in (today.year, today.year + 1):
        try:
            candidate = dt.date(year, month, day)
        except ValueError:
            continue
        if candidate >= today:
            return candidate
    return None


def normalise_eta_date(text):
    """Latest delivery date in ISO form, or None when it cannot be determined
    confidently. Returning None is always preferred over a guess - the caller
    keeps the raw wording either way."""
    if not text:
        return None

    candidates = []

    for day_a, day_b, month_name in _BARE_RANGE_RE.findall(text):
        month = _month_number(month_name)
        if month:
            for day in (day_a, day_b):
                resolved = _resolve_future_date(int(day), month)
                if resolved:
                    candidates.append(resolved)

    for day, month_name in _DATE_DM_RE.findall(text):
        month = _month_number(month_name)
        if month:
            resolved = _resolve_future_date(int(day), month)
            if resolved:
                candidates.append(resolved)

    for month_name, day in _DATE_MD_RE.findall(text):
        month = _month_number(month_name)
        if month:
            resolved = _resolve_future_date(int(day), month)
            if resolved:
                candidates.append(resolved)

    if not candidates:
        return None
    return max(candidates).isoformat()


_AVAILABILITY_OUT_OF_STOCK_PHRASES = ("out of stock", "unavailable", "currently unavailable")
_AVAILABILITY_IN_STOCK_PHRASES = ("in stock",)

# id="availability" is the normal container. On a genuinely out-of-stock real
# Amazon AU page (confirmed 2026-08-05 against a live-fetched, cached page -
# confirmed against a live-fetched page),
# id="availability" is absent entirely and the unavailable notice lives in
# id="outOfStock" / id="outOfStockBuyBox_feature_div" instead. Checked in order;
# first one present with a recognizable phrase wins.
_AVAILABILITY_ELEMENT_IDS = ("availability", "outOfStock", "outOfStockBuyBox_feature_div")


def parse_availability(html_text):
    for element_id in _AVAILABILITY_ELEMENT_IDS:
        text = extract_element_text_by_id(html_text, element_id)
        if not text:
            continue
        low = text.lower()
        if any(phrase in low for phrase in _AVAILABILITY_OUT_OF_STOCK_PHRASES):
            return "OUT_OF_STOCK"
        if any(phrase in low for phrase in _AVAILABILITY_IN_STOCK_PHRASES):
            return "IN_STOCK"
        if "left" in low and re.search(r"\bonly\b", low):
            return "LIMITED"
    return None


_GLOW_LOCATION_RE = re.compile(
    r'id=["\']glow-ingress-line2["\'][^>]*>(.*?)</', re.IGNORECASE | re.DOTALL)
_DELIVER_TO_RE = re.compile(r'Deliver to ([A-Za-z][A-Za-z .\-]{1,40})')


def extract_delivery_location(html_text):
    """The delivery location Amazon ACTUALLY used to compute this page's offers,
    ETA, and fees - read from its own location widget, not assumed.

    This matters more than it looks. ETA and delivery fee are properties of a
    viewer in a place (PLAN.md section 8). Logged out with no postcode set,
    Amazon geolocates from the exit IP, which is not necessarily Australia. A
    row stamped "AU postcode 2000" that was actually measured from somewhere
    else is worse than no row at all - RC would set real shipping policy from
    it. Returns a location string, or None if it can't be determined.
    """
    match = _GLOW_LOCATION_RE.search(html_text)
    if match:
        text = _strip_tags(match.group(1))
        if text:
            return text
    match = _DELIVER_TO_RE.search(html_text)
    if match:
        return match.group(1).strip()
    return None


def parse_fixture(path, requested_asin, log):
    """Parse one saved Amazon AU HTML fixture, offline. Returns a result dict
    shaped like a MOCK_RESULTS entry (status/reason plus whichever fields were
    found). Never raises - any unexpected failure is caught and routed to
    errors, so one bad saved page can't take down the whole run.
    """
    try:
        html_text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as exc:
        return {"status": "parse_error", "reason": "fixture_unreadable:%s" % exc}
    return parse_html(html_text, requested_asin, log)


def parse_html(html_text, requested_asin, log):
    """Parse one Amazon AU page from an in-memory string.

    Split out from parse_fixture() so a live fetch can be parsed without ever
    touching the disk - production discards HTML after this returns. Parsing
    behaviour is unchanged; parse_fixture() is now just "read file, call this".
    """
    try:
        canonical = extract_canonical_asin(html_text)
        if canonical and canonical != requested_asin:
            return {
                "status": "redirect",
                "reason": "canonical_asin_differs",
                "final_asin": canonical,
            }

        sold_by, ships_from = parse_merchant_fields(html_text)
        delivery = parse_delivery(html_text)
        availability = parse_availability(html_text)

        result = {}

        # Record the basis actually observed, never the assumed one. See
        # extract_delivery_location() for why this must not be hardcoded.
        observed_location = extract_delivery_location(html_text)
        if observed_location:
            result["delivery_location_basis"] = "observed: %s (logged-out, non-Prime)" % observed_location
            if not re.search(r"australia|\bAU\b|\b\d{4}\b", observed_location, re.IGNORECASE):
                log("WARNING: %s delivery basis is %r, NOT Australia - any ETA/fee on this "
                    "row is measured from the wrong country" % (requested_asin, observed_location))
        else:
            result["delivery_location_basis"] = "UNVERIFIED assumed: %s" % DELIVERY_LOCATION_BASIS
        if sold_by:
            # merchant == sold_by and delivered_by == ships_from for now - open
            # question in PLAN.md section 7, not yet confirmed with RC.
            result["sold_by"] = sold_by
            result["merchant"] = sold_by
        if ships_from:
            result["ships_from"] = ships_from
            result["shipped_by"] = ships_from
            result["delivered_by"] = ships_from
        result.update(delivery)
        if availability:
            result["availability"] = availability

        found_shipping_data = bool(sold_by or ships_from or delivery.get("delivery_message_raw"))
        if not found_shipping_data:
            if availability == "OUT_OF_STOCK":
                # A genuinely unavailable item has no active offer to show
                # seller/delivery info for - that's a real answer, not a gap.
                # See PLAN.md section 10: out_of_stock is a legitimate finding,
                # not a failure, and distinct from "we don't know".
                result["status"] = "out_of_stock"
                result["reason"] = "out_of_stock_no_active_offer"
            else:
                result["status"] = "no_delivery_info"
                result["reason"] = "no_shipping_merchant_data_found_in_fixture"
        elif "delivery_fee" not in result:
            result["status"] = "no_delivery_fee"
            result["reason"] = "delivery_block_present_no_fee_parsed"
        else:
            result["status"] = "extracted"
            result["reason"] = ""
        return result
    except Exception as exc:  # conservative: never let one bad page crash the run
        log("fixture parse error for %s: %r" % (path, exc))
        return {"status": "parse_error", "reason": "unexpected_parse_exception:%s" % type(exc).__name__}


# --------------------------------------------------------------------------
# Live fetch - tiny direct pilot only, no proxy, no Decodo, no concurrency
#
# This is the ONLY place in the file that can make a network call, and it only
# runs when --live is passed. One request at a time, from a single-threaded
# loop, with a mandatory sleep between requests and a hard --limit on how many
# fetches a single run will make. No proxy handler is ever attached, and an
# explicit no-proxy opener is used so a stray http_proxy/https_proxy environment
# variable can't smuggle one in.
# --------------------------------------------------------------------------

_LIVE_REQUEST_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "en-AU,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    # urllib does NOT send this by default, so Amazon was returning the product
    # page UNCOMPRESSED. Measured on a real page: 1,686,090 wire bytes without
    # it vs 345,466 with it - a 4.9x (80%) reduction for identical content.
    # This is the single largest data saving in the tool.
    "Accept-Encoding": "gzip, deflate",
}

# Explicitly no proxy, regardless of http_proxy/https_proxy in the environment.
# The cookie jar is what carries the Australian delivery context: Amazon sets a
# location cookie once (see establish_au_delivery_context) and every later
# product request reuses it, so the AU context costs ONE extra request per run,
# not one per ASIN.
_COOKIE_JAR = http.cookiejar.CookieJar()

# Default: no proxy at all, regardless of http_proxy/https_proxy in the env.
# configure_proxy() swaps this for a proxied opener when one is supplied.
_NO_PROXY_OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({}),
    urllib.request.HTTPCookieProcessor(_COOKIE_JAR),
)

# Set once at startup. Kept module-level so every request path (AU location
# setup and product fetches alike) goes through the same proxy and the same
# cookie jar - a split would put the delivery-location cookie on one exit IP
# and the product request on another, which Amazon treats as a new visitor.
PROXY_STATE = {"configured": False, "host": None}


def _scrub_proxy_secrets(text):
    """Never let proxy credentials reach a log, a CSV, or the console."""
    return re.sub(r"//[^/@\s]*:[^/@\s]*@", "//***:***@", str(text))


def configure_proxy(proxy_url, log):
    """Route all HTTP through proxy_url (http://user:pass@host:port).

    Amazon throttles by IP reputation - a datacenter/residential-flagged address
    gets served a 'Server Busy' stub instead of the product page. A residential
    AU exit both lifts that and puts the delivery context in Australia.

    Returns True when a proxy is active. Credentials are read from the
    environment or CLI and are never written to disk or logged in the clear.
    """
    global _NO_PROXY_OPENER
    if not proxy_url:
        return False
    parsed = urllib.parse.urlsplit(proxy_url)
    if not parsed.hostname:
        raise SystemExit("--proxy must look like http://user:pass@host:port (got %r)"
                         % _scrub_proxy_secrets(proxy_url))
    _NO_PROXY_OPENER = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url}),
        urllib.request.HTTPCookieProcessor(_COOKIE_JAR),
    )
    PROXY_STATE["configured"] = True
    PROXY_STATE["host"] = "%s:%s" % (parsed.hostname, parsed.port or "?")
    log("proxy ENABLED via %s (credentials hidden)" % PROXY_STATE["host"])
    return True

_GLOW_TOKEN_RE = re.compile(
    r'id="glowValidationToken"[^>]*value="([^"]+)"|'
    r'name="glow-validation-token"[^>]*value="([^"]+)"', re.IGNORECASE)


def establish_au_delivery_context(postcode, timeout_s, log):
    """Set the session's delivery location to an Australian postcode, ONCE per run.

    Why this exists: amazon.com.au geolocates a logged-out visitor from their exit
    IP. Running from outside Australia, Amazon renders a degraded page - live
    bestsellers report "Currently unavailable", and ETA/fee/merchant are absent.
    That is not real stock data, and parsing it would feed RC wrong shipping info.

    This uses Amazon's own logged-out "Deliver to" widget - the same form
    submission the site's JavaScript makes when a visitor types a postcode,
    including the validation token the page itself hands out for it. No login, no
    CAPTCHA handling, no access control is involved.

    Returns (ok, wire_bytes). On failure the caller continues; every row then
    records the real (non-AU) location rather than pretending it worked.
    """
    # Seed page choice is purely about cost. The location widget's token appears
    # on full store pages, not on the bare homepage (which returns a ~1 KB stub).
    # A category page carries it for ~111 KB, vs ~400 KB for a product page -
    # measured, and paid once per run.
    seed_url = "https://www.amazon.com.au/gp/bestsellers/kitchen/"
    try:
        seed_req = urllib.request.Request(seed_url, headers=_LIVE_REQUEST_HEADERS)
        response = _NO_PROXY_OPENER.open(seed_req, timeout=timeout_s)
        raw = response.read()
        response.close()
        seed_bytes = len(raw)
        if (response.headers.get("content-encoding") or "").lower() == "gzip":
            raw = gzip.decompress(raw)
        seed_html = raw.decode("utf-8", errors="replace")
    except Exception as exc:
        log("AU delivery context: seed request failed (%s) - continuing without it"
            % type(exc).__name__)
        return False, 0

    match = _GLOW_TOKEN_RE.search(seed_html)
    token = (match.group(1) or match.group(2)) if match else None
    if not token:
        log("AU delivery context: page did not offer a location-widget token - "
            "continuing without it")
        return False, seed_bytes

    payload = urllib.parse.urlencode({
        "locationType": "LOCATION_INPUT",
        "zipCode": postcode,
        "storeContext": "generic",
        "deviceType": "web",
        "pageType": "Detail",
        "actionSource": "glow",
    }).encode()
    headers = dict(_LIVE_REQUEST_HEADERS)
    headers.update({
        "anti-csrftoken-a2z": token,
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": seed_url,
    })
    try:
        req = urllib.request.Request(
            "https://www.amazon.com.au/portal-migration/hz/glow/address-change?actionSource=glow",
            data=payload, headers=headers)
        response = _NO_PROXY_OPENER.open(req, timeout=timeout_s)
        post_bytes = len(response.read())
        status = response.getcode()
        response.close()
    except Exception as exc:
        log("AU delivery context: location request failed (%s) - continuing without it"
            % type(exc).__name__)
        return False, seed_bytes

    total = seed_bytes + post_bytes
    if status != 200:
        log("AU delivery context: location request returned HTTP %d - continuing without it" % status)
        return False, total
    log("AU delivery context established for postcode %s (2 setup requests, %.1f KB total, "
        "reused for every ASIN this run)" % (postcode, total / 1024.0))
    return True, total

_BLOCK_SIGNAL_PHRASES = (
    "enter the characters you see below",
    "to discuss automated access to amazon data",
    "sorry, we just need to make sure you're not a robot",
    "api.amazon.com/captcha",
)
_TITLE_RE = re.compile(r"<title[^>]*>(.*?)</title>", re.IGNORECASE | re.DOTALL)


# Bytes actually transferred by the most recent fetch, so per-ASIN data usage is
# measured rather than estimated.
LAST_FETCH_STATS = {"wire_bytes": 0, "decoded_bytes": 0}


def _short_reason(value):
    text = str(value) if value else ""
    return text[:80] if text else "unknown"


def fetch_live_html(url, timeout_s):
    """Single direct GET, no retries. Returns (html_text, error) where exactly
    one is None. error is a short machine-readable string for the errors-sheet
    reason column: "http_status_<code>", "proxy_auth_failed", or
    "fetch_failed:<reason>".
    """
    request = urllib.request.Request(url, headers=_LIVE_REQUEST_HEADERS)
    try:
        response = _NO_PROXY_OPENER.open(request, timeout=timeout_s)
    except urllib.error.HTTPError as exc:
        # 407 means the proxy rejected our credentials. Retrying cannot help,
        # and a retry loop on this exact fault is what turned an outage into
        # ~263k wasted proxy connections in the previous system. Surface it as
        # its own status so the run can stop instead of grinding.
        if exc.code == 407:
            return None, "proxy_auth_failed"
        return None, "http_status_%d" % exc.code
    except urllib.error.URLError as exc:
        reason = _scrub_proxy_secrets(_short_reason(exc.reason))
        if "407" in reason or "proxy" in reason.lower():
            return None, "proxy_auth_failed"
        return None, "fetch_failed:%s" % reason
    except Exception as exc:
        return None, "fetch_failed:%s" % type(exc).__name__

    try:
        status = response.getcode()
        if status != 200:
            return None, "http_status_%d" % status
        charset = response.headers.get_content_charset() or "utf-8"
        encoding = (response.headers.get("content-encoding") or "").lower()
        body = response.read()
    except Exception as exc:
        return None, "fetch_failed:%s" % type(exc).__name__
    finally:
        response.close()

    # Record what actually crossed the wire, before decompression - this is the
    # number that maps to bandwidth cost, not len(text).
    LAST_FETCH_STATS["wire_bytes"] = len(body)

    try:
        if encoding == "gzip":
            body = gzip.decompress(body)
        elif encoding == "deflate":
            try:
                body = zlib.decompress(body)
            except zlib.error:
                body = zlib.decompress(body, -zlib.MAX_WBITS)
    except Exception as exc:
        return None, "fetch_failed:decompress_%s" % type(exc).__name__

    LAST_FETCH_STATS["decoded_bytes"] = len(body)

    try:
        text = body.decode(charset, errors="replace")
    except LookupError:
        text = body.decode("utf-8", errors="replace")
    return text, None


def _looks_blocked_or_busy(html_text):
    """Conservative, self-authored heuristic - fires only on well-known Amazon
    block/CAPTCHA/throttle text, never guesses. Does not import WEB-SCRAPER."""
    low = html_text.lower()
    if any(phrase in low for phrase in _BLOCK_SIGNAL_PHRASES):
        return True
    title_match = _TITLE_RE.search(html_text)
    if title_match:
        title_text = _strip_tags(title_match.group(1)).lower()
        if title_text == "server busy" or title_text.startswith("server busy"):
            return True
    return False


# --------------------------------------------------------------------------
# Browser-rendered fetch - tiny direct pilot only, no proxy, no Decodo, no
# concurrency, no package install
#
# Real seller/delivery data was confirmed (2026-08-05, against a live-fetched,
# cached, real Amazon AU page) to be JS-hydrated: the feature names
# ("merchantInfoFeature_feature_div" etc.) exist only as string keys inside an
# embedded JS lazy-load registry, never as populated DOM elements, in a plain
# GET response. A raw fetch structurally cannot see this data - only a browser
# that executes JavaScript can. See
#
# Neither Playwright nor Selenium is installed in this environment (checked,
# not installed silently). Rather than add a new dependency, this uses Chrome's
# own built-in headless CLI mode (--headless=new --dump-dom), which is already
# present as a side effect of Chrome.app being installed - zero packages added.
# --------------------------------------------------------------------------

_CHROME_CANDIDATE_PATHS = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Google Chrome Canary.app/Contents/MacOS/Google Chrome Canary",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
)


def find_chrome_binary():
    """First existing path from a short list of common Chrome/Chromium
    locations, or None. Never installs anything."""
    for path in _CHROME_CANDIDATE_PATHS:
        if os.path.exists(path):
            return path
    return None


_CANONICAL_URL_RE = re.compile(
    r'<link[^>]+rel=["\']canonical["\'][^>]*href=["\']([^"\']+)["\']', re.IGNORECASE)
_OG_URL_RE = re.compile(
    r'<meta[^>]+property=["\']og:url["\'][^>]*content=["\']([^"\']+)["\']', re.IGNORECASE)
_HOST_RE = re.compile(r'^https?://([^/]+)', re.IGNORECASE)
EXPECTED_MARKETPLACE_HOSTS = ("amazon.com.au", "www.amazon.com.au")


def extract_served_host(html_text):
    """Which marketplace actually served this page, per the page's own canonical
    /og:url tag. Returns a hostname or None.

    NOTE ON WHAT THIS IS: Chrome's --dump-dom prints the DOM, not the address
    bar, so this is Amazon's own declaration of the page's marketplace rather
    than a literal readout of the browser's final URL after redirects. For
    "did we land on AU or get bounced to another marketplace" that is the
    meaningful signal - a US-served page carries an amazon.com canonical - but
    it is inferred from page content, not observed from the browser.
    """
    for pattern in (_CANONICAL_URL_RE, _OG_URL_RE):
        match = pattern.search(html_text)
        if match:
            host_match = _HOST_RE.match(match.group(1).strip())
            if host_match:
                return host_match.group(1).lower()
    return None


def fetch_rendered_html(chrome_binary, url, timeout_s):
    """Single headless-Chrome DOM dump via subprocess. No proxy
    (--no-proxy-server, explicit regardless of environment), no extensions, one
    process at a time - called at most once per ASIN per run, from the same
    single-threaded loop as fetch_live_html(), with a sleep between calls.
    --virtual-time-budget gives async/JS-loaded content a window to complete
    before the DOM is captured. Returns (html_text, error); exactly one is None.
    """
    cmd = [
        chrome_binary,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-proxy-server",
        "--disable-extensions",
        "--lang=en-AU",
        "--user-agent=%s" % _LIVE_REQUEST_HEADERS["User-Agent"],
        # Only load what is needed to read merchant / delivery text out of the
        # DOM. A product page otherwise pulls megabytes of photography, fonts,
        # video and tracking beacons we never look at.
        "--blink-settings=imagesEnabled=false",
        "--disable-remote-fonts",
        "--autoplay-policy=user-gesture-required",
        "--mute-audio",
        "--disable-background-networking",
        "--disable-sync",
        "--disable-default-apps",
        "--no-first-run",
        "--disable-client-side-phishing-detection",
        "--disable-component-update",
        "--disable-domain-reliability",
        "--metrics-recording-only",
        "--no-pings",
        "--virtual-time-budget=8000",
        "--dump-dom",
        url,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout_s)
    except subprocess.TimeoutExpired:
        return None, "fetch_failed:browser_timeout"
    except Exception as exc:
        return None, "fetch_failed:%s" % type(exc).__name__

    if proc.returncode != 0:
        return None, "fetch_failed:browser_exit_%d" % proc.returncode
    html_text = proc.stdout
    if not html_text or len(html_text.strip()) < 50:
        return None, "fetch_failed:browser_empty_dom"
    return html_text, None


# --------------------------------------------------------------------------
# HTML retention policy
#
# Production discards page HTML as soon as an ASIN is parsed. At 200k ASINs a
# ~2.4 MB page would be ~480 GB of disk if kept, and a cached page is a stale
# snapshot - delivery ETAs change daily, so silently reusing one would report
# yesterday's ETA as today's. Two narrow exceptions:
#   * --debug-cache        - full dev behaviour, opt-in, unbounded (dev only)
#   * bounded failure cache - only genuine anomalies, hard file cap, oldest
#                             rotated out, so it cannot grow without limit.
# --------------------------------------------------------------------------

# Statuses worth keeping HTML for: the page was unexpected or the parser could
# not read it. Ordinary business outcomes (out of stock, redirect, no fee found)
# are NOT anomalies and are never cached.
_FAILURE_CACHE_STATUSES = frozenset({
    "parse_error", "unsupported_layout", "server_busy",
})


def save_failure_html(failure_dir, asin, html_text, status, max_files, log):
    """Keep a bounded sample of anomalous pages for debugging.

    Rotates: once the cap is hit, the oldest file is deleted before writing a
    new one, so disk use is capped at roughly max_files x page size regardless
    of how long the batch runs.
    """
    if not max_files or failure_dir is None or not html_text:
        return False
    try:
        failure_dir.mkdir(parents=True, exist_ok=True)
        existing = sorted(failure_dir.glob("*.html"), key=lambda p: p.stat().st_mtime)
        while len(existing) >= max_files:
            oldest = existing.pop(0)
            try:
                oldest.unlink()
            except OSError:
                break
        (failure_dir / ("%s_%s.html" % (status, asin))).write_text(html_text, encoding="utf-8")
        log("  saved failure HTML for %s (%s), cap %d" % (asin, status, max_files))
        return True
    except OSError as exc:
        log("  could not save failure HTML for %s: %s" % (asin, type(exc).__name__))
        return False


def _retain_html(html_text, asin, result, cache_dir, retention, log):
    """Apply the retention policy to one page. Default: keep nothing."""
    if retention is None:
        return
    if retention.get("debug_cache") and cache_dir is not None:
        cache_dir.mkdir(parents=True, exist_ok=True)
        (cache_dir / ("%s.html" % asin)).write_text(html_text, encoding="utf-8")
        log("  [debug-cache] wrote %s.html" % asin)
        return
    status = result.get("status", "")
    if status in _FAILURE_CACHE_STATUSES:
        if save_failure_html(retention.get("failure_dir"), asin, html_text, status,
                             retention.get("failure_max", 0), log):
            retention["failures_saved"] = retention.get("failures_saved", 0) + 1
    # Anything else: HTML is simply not written. It goes out of scope with the
    # caller's frame and is reclaimed normally.


class RunAborted(Exception):
    """A safety limit tripped. The batch stops; results so far are still written."""


def _trip_if_unhealthy(live_state, live_cfg, log):
    """Stop the batch when the run is clearly not working.

    Without this a broken proxy or a block wave quietly burns the whole ASIN
    list - the exact failure mode that ran for 10+ hours in the previous system.
    Fail closed: stop and let a human look, never auto-resume.
    """
    done = live_state["fetches_done"]
    blocked = live_state["blocked"]
    auth_failed = live_state["proxy_auth_failures"]

    if auth_failed >= live_cfg["max_proxy_auth_failures"]:
        raise RunAborted("proxy rejected credentials %d times - check the proxy user/password"
                         % auth_failed)

    # Only judge the block rate once there is enough signal to judge it on.
    if done >= live_cfg["block_rate_min_samples"]:
        rate = blocked / float(done)
        if rate >= live_cfg["max_block_rate"]:
            raise RunAborted(
                "%.0f%% of the last %d fetches were blocked by Amazon (limit %.0f%%) - "
                "stopping so the rest of the list isn't wasted"
                % (rate * 100, done, live_cfg["max_block_rate"] * 100))


def classify(unit_key, fixtures_dir, cache_dir, live_cfg, live_state,
             rendered_cache_dir, browser_cfg, browser_state, log, retention=None):
    """Classification for one fetch unit (ASIN or SKU key).

    Priority:
      1. rendered_cache/<ASIN>.html - always checked first when present; a
         rendered page is a superset of a raw one, so it wins regardless of
         which mode this run is using.
      2. page_cache/<ASIN>.html - reused and reparsed, never refetched.
      3. --browser-live only, no cache: one headless-Chrome DOM dump (subject
         to --browser-limit/--browser-delay-seconds), cached, then parsed.
      4. --live only, no cache: one direct raw fetch (subject to --limit/
         --delay-seconds), cached, then parsed.
      5. Otherwise: the Phase 2 fixtures/ lookup, then the mock table, then a
         clear "nothing available" marker.
    fixtures_dir/cache_dir/rendered_cache_dir may be None (--mock-only),
    reproducing Phase 1 behavior exactly. live_cfg/browser_cfg are None unless
    the matching flag was passed; only one of --live/--browser-live/--mock-only
    can be active in a given run (enforced in main()).
    """
    is_asin = bool(ASIN_RE.match(unit_key))

    # Cached HTML is a snapshot: delivery ETAs change daily, so reusing one would
    # report a stale ETA as current. Reads are therefore OPT-IN, never automatic.
    read_cache = bool(retention and retention.get("read_cache"))

    if read_cache and rendered_cache_dir is not None and is_asin:
        rendered_path = rendered_cache_dir / ("%s.html" % unit_key)
        if rendered_path.exists():
            result = parse_fixture(rendered_path, unit_key, log)
            log("rendered-cache classify %s -> %s (%s)" % (unit_key, result.get("status"), rendered_path.name))
            return result

    if read_cache and cache_dir is not None and is_asin:
        cache_path = cache_dir / ("%s.html" % unit_key)
        if cache_path.exists():
            result = parse_fixture(cache_path, unit_key, log)
            log("cache classify %s -> %s (%s)" % (unit_key, result.get("status"), cache_path.name))
            return result

    if browser_cfg is not None:
        if not is_asin:
            log("no cache and --browser-live set, but %r is not ASIN-shaped" % unit_key)
            return {"status": "no_delivery_info", "reason": "browser_fetch_requires_asin"}

        if browser_state["fetches_done"] >= browser_cfg["limit"]:
            log("browser fetch limit (%d) reached - skipping %s" % (browser_cfg["limit"], unit_key))
            return {"status": "no_delivery_info", "reason": "browser_fetch_limit_reached"}

        if browser_state["fetches_done"] > 0:
            log("sleeping %.1fs before next browser fetch" % browser_cfg["delay_seconds"])
            time.sleep(browser_cfg["delay_seconds"])

        url = "https://www.amazon.com.au/dp/%s" % unit_key
        log("BROWSER fetch %s -> %s" % (unit_key, url))
        html_text, error = fetch_rendered_html(browser_cfg["chrome_binary"], url, browser_cfg["timeout_seconds"])
        browser_state["fetches_done"] += 1

        if error is not None:
            log("browser fetch failed for %s: %s" % (unit_key, error))
            return {"status": "fetch_failed", "reason": error}

        if _looks_blocked_or_busy(html_text):
            log("browser fetch for %s looks blocked/captcha/Server Busy" % unit_key)
            result = {"status": "server_busy", "reason": "blocked_or_captcha_or_server_busy"}
            _retain_html(html_text, unit_key, result, rendered_cache_dir, retention, log)
            return result

        # Marketplace guard: we requested amazon.com.au, so the page that came
        # back must be an AU page. A US/other-marketplace page would have
        # different sellers, prices, and delivery entirely - parsing it as if it
        # were AU would silently produce wrong shipping data for RC.
        served_host = extract_served_host(html_text)
        browser_state["served_hosts"][unit_key] = served_host
        if served_host is None:
            log("WARNING: could not determine served marketplace host for %s" % unit_key)
        elif served_host not in EXPECTED_MARKETPLACE_HOSTS:
            log("STOP: %s served by %r, not Amazon AU - refusing to parse as AU data"
                % (unit_key, served_host))
            return {"status": "fetch_failed",
                    "reason": "wrong_marketplace_host:%s" % served_host}
        else:
            log("marketplace confirmed for %s: %s" % (unit_key, served_host))

        result = parse_html(html_text, unit_key, log)
        log("browser classify %s -> %s" % (unit_key, result.get("status")))
        _retain_html(html_text, unit_key, result, rendered_cache_dir, retention, log)
        return result

    if live_cfg is not None:
        if not is_asin:
            log("no cache and --live set, but %r is not ASIN-shaped - nothing to fetch" % unit_key)
            return {"status": "no_delivery_info", "reason": "live_fetch_requires_asin"}

        if live_state["fetches_done"] >= live_cfg["limit"]:
            log("live fetch limit (%d) reached - skipping %s" % (live_cfg["limit"], unit_key))
            return {"status": "no_delivery_info", "reason": "live_fetch_limit_reached"}

        if live_state["fetches_done"] > 0:
            log("sleeping %.1fs before next live fetch" % live_cfg["delay_seconds"])
            time.sleep(live_cfg["delay_seconds"])

        url = "https://www.amazon.com.au/dp/%s" % unit_key
        log("LIVE fetch %s -> %s" % (unit_key, url))
        LAST_FETCH_STATS.update({"wire_bytes": 0, "decoded_bytes": 0})
        html_text, error = fetch_live_html(url, live_cfg["timeout_seconds"])
        live_state["fetches_done"] += 1
        live_state["wire_bytes"] += LAST_FETCH_STATS["wire_bytes"]
        log("  transferred %.0f KB on the wire (%.2f MB decompressed)"
            % (LAST_FETCH_STATS["wire_bytes"] / 1024.0,
               LAST_FETCH_STATS["decoded_bytes"] / 1048576.0))

        if error is not None:
            log("live fetch failed for %s: %s" % (unit_key, error))
            if error == "proxy_auth_failed":
                live_state["proxy_auth_failures"] += 1
                _trip_if_unhealthy(live_state, live_cfg, log)
                return {"status": "proxy_auth_failed", "reason": error}
            status = "not_found" if error == "http_status_404" else "fetch_failed"
            return {"status": status, "reason": error}

        if _looks_blocked_or_busy(html_text):
            log("live fetch for %s looks blocked/captcha/Server Busy" % unit_key)
            live_state["blocked"] += 1
            _trip_if_unhealthy(live_state, live_cfg, log)
            result = {"status": "server_busy", "reason": "blocked_or_captcha_or_server_busy"}
        else:
            # Parse straight from memory. The HTML is never written to disk on
            # the success path and is released when this frame returns.
            result = parse_html(html_text, unit_key, log)
            log("live classify %s -> %s" % (unit_key, result.get("status")))

        _retain_html(html_text, unit_key, result, cache_dir, retention, log)
        return result

    if fixtures_dir is not None and is_asin:
        fixture_path = fixtures_dir / ("%s.html" % unit_key)
        if fixture_path.exists():
            result = parse_fixture(fixture_path, unit_key, log)
            log("fixture classify %s -> %s (%s)" % (unit_key, result.get("status"), fixture_path.name))
            return result

    if unit_key in MOCK_RESULTS:
        result = MOCK_RESULTS[unit_key]
        log("mock classify %s -> %s (no fixture found)" % (unit_key, result["status"]))
        return result

    log("no fixture and no mock result for %s" % unit_key)
    return DEFAULT_MOCK


def project_to_output(row, result, extracted_at):
    """Map an internal parse result onto the slim output contract.

    The parser keeps richer internal keys (sold_by, ships_from, availability...)
    because they drive classification, but only the contracted COLUMNS are ever
    written out. This is the single place that decides what leaves the tool.
    """
    rec = _blank_record()
    rec["sku"] = row["sku"]
    rec["asin"] = row["asin"]
    rec["checked_at"] = extracted_at
    rec["status"] = result.get("status", "")
    rec["error"] = result.get("reason", "")
    rec["eta_raw"] = result.get("delivery_message_raw", "")
    rec["eta_date"] = result.get("eta_date", "")
    # One merchant field, from the buy-box offer we already have. No seller
    # profile page is ever fetched to enrich this.
    rec["merchant"] = result.get("sold_by", "") or result.get("merchant", "")
    fee = result.get("delivery_fee", "")
    rec["delivery_fee"] = fee if fee != "" else ""
    rec["delivery_location"] = result.get("delivery_location_basis", "")
    # A redirect is still a real outcome for this row; surface where it went in
    # the error column rather than adding a column nobody asked for.
    final_asin = result.get("final_asin")
    if final_asin and final_asin != row["asin"]:
        rec["error"] = ("%s->%s" % (rec["error"], final_asin)).lstrip("-")
    return rec


def load_completed_asins(run_dir, log):
    """ASINs already finished in an earlier run, read from its progress.jsonl.

    A truncated final line (process killed mid-write) is skipped rather than
    treated as corruption - that record simply gets redone.
    """
    path = os.path.join(run_dir, "progress.jsonl")
    if not os.path.exists(path):
        raise SystemExit("--resume: no progress.jsonl in %s" % run_dir)
    done = set()
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                done.add(json.loads(line)["asin"])
            except (ValueError, KeyError):
                continue
    log("--resume: %d ASINs already completed in %s" % (len(done), run_dir))
    return done


def build_records(units, rejected, extracted_at, fixtures_dir, cache_dir, live_cfg, live_state,
                   rendered_cache_dir, browser_cfg, browser_state, log,
                   retention=None, progress_path=None):
    """Turn fetch units and rejected rows into one record per input row.

    Each completed fetch unit is appended to progress.jsonl immediately, so a
    crash loses at most the ASIN in flight. This is what makes discarding page
    HTML safe at scale - resume no longer depends on a page cache.
    """
    records = []
    order = {}

    def checkpoint(asin, result):
        if not progress_path:
            return
        try:
            with open(progress_path, "a", encoding="utf-8") as fh:
                fh.write(json.dumps({"asin": asin, "status": result.get("status", ""),
                                     "at": dt.datetime.now().isoformat(timespec="seconds")}) + "\n")
        except OSError:
            pass  # never let checkpoint I/O kill an in-flight batch

    for row in rejected:
        rec = project_to_output(
            row, {"status": "invalid_row", "reason": row["reject_reason"]}, extracted_at)
        order[id(rec)] = row["input_row_number"]
        records.append(rec)

    for key, member_rows in units.items():
        try:
            result = classify(key, fixtures_dir, cache_dir, live_cfg, live_state,
                              rendered_cache_dir, browser_cfg, browser_state, log, retention)
        except RunAborted as exc:
            # Stop fetching, but keep everything already extracted: the partial
            # results are written normally and progress.jsonl lets --resume pick
            # up from here once the cause is fixed.
            log("ABORTED: %s" % exc)
            log("stopping after %d fetches - results so far are still written"
                % live_state.get("fetches_done", 0))
            live_state["aborted"] = str(exc)
            break
        for row in member_rows:
            rec = project_to_output(row, result, extracted_at)
            order[id(rec)] = row["input_row_number"]
            records.append(rec)
        # A unit skipped only because --limit was hit was never actually
        # fetched - it must stay eligible for a future --resume, not get
        # checkpointed as done.
        if result.get("reason") not in ("live_fetch_limit_reached", "browser_fetch_limit_reached"):
            checkpoint(key, result)

    records.sort(key=lambda r: order[id(r)])
    return records


# --------------------------------------------------------------------------
# Excel writing
# --------------------------------------------------------------------------

def _col_letter(index):
    """1-based column index -> spreadsheet column letter."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _write_xlsx_stdlib(path, sheets):
    """Minimal XLSX writer using only the standard library.

    SCAFFOLDING: exists so Phase 1 could be validated without installing openpyxl.
    Delete once openpyxl is available - write_xlsx already prefers it.
    """
    import zipfile
    from xml.sax.saxutils import escape

    def cell_xml(ref, value):
        if value is None or value == "":
            return ""
        if isinstance(value, bool):
            return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, "TRUE" if value else "FALSE")
        if isinstance(value, (int, float)):
            return '<c r="%s"><v>%s</v></c>' % (ref, value)
        return '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, escape(str(value)))

    def sheet_xml(rows):
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                 '<sheetData>']
        for r_idx, row in enumerate(rows, start=1):
            cells = "".join(cell_xml("%s%d" % (_col_letter(c_idx), r_idx), value)
                            for c_idx, value in enumerate(row, start=1))
            parts.append('<row r="%d">%s</row>' % (r_idx, cells))
        parts.append("</sheetData></worksheet>")
        return "".join(parts)

    n = len(sheets)
    styles_rid = n + 1

    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(1, n + 1):
        content_types.append('<Override PartName="/xl/worksheets/sheet%d.xml" '
                             'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    content_types.append("</Types>")

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                 'Target="xl/workbook.xml"/></Relationships>')

    sheet_tags = "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(name), i, i)
                         for i, (name, _) in enumerate(sheets, start=1))
    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>%s</sheets></workbook>' % sheet_tags)

    rel_items = "".join(
        '<Relationship Id="rId%d" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet%d.xml"/>' % (i, i) for i in range(1, n + 1))
    rel_items += ('<Relationship Id="rId%d" '
                  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                  'Target="styles.xml"/>' % styles_rid)
    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '%s</Relationships>' % rel_items)

    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border/></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
              '</styleSheet>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles)
        for i, (_, rows) in enumerate(sheets, start=1):
            zf.writestr("xl/worksheets/sheet%d.xml" % i, sheet_xml(rows))


def write_xlsx(path, sheets, log):
    """Write the workbook. Prefers openpyxl; falls back to the stdlib writer."""
    try:
        from openpyxl import Workbook
    except ImportError:
        log("openpyxl not installed - using the standard-library XLSX fallback")
        _write_xlsx_stdlib(path, sheets)
        return "stdlib-fallback"

    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets:
        ws = wb.create_sheet(title=name)
        for row in rows:
            ws.append(list(row))
    wb.save(path)
    log("wrote workbook with openpyxl")
    return "openpyxl"


# --------------------------------------------------------------------------
# Run
# --------------------------------------------------------------------------

def build_summary_rows(summary):
    rows = [["metric", "value"]]
    for key, value in summary.items():
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                rows.append(["%s.%s" % (key, sub_key), sub_value])
        elif isinstance(value, (list, tuple, set)):
            rows.append([key, "; ".join(str(v) for v in value)])
        else:
            rows.append([key, value])
    return rows


def write_csv_rows(path, header, data_rows, log):
    """Write one CSV review file. UTF-8, stable column order, header first row.

    Values are taken as-is from the same records that feed the workbook sheets, so
    blank-vs-zero (e.g. unknown delivery fee vs FREE delivery) can never diverge
    between the XLSX and CSV outputs - there is exactly one source of truth per row.
    """
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(header)
        writer.writerows(data_rows)
    log("wrote %s" % path)


def unique_path(path):
    """Never overwrite an existing workbook."""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 2
    while os.path.exists("%s_%d%s" % (stem, n, ext)):
        n += 1
    return "%s_%d%s" % (stem, n, ext)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Shipping Enrichment Lite - Phase 1 + 2 + tiny live/browser pilot (no proxy, no Decodo)")
    parser.add_argument("--input", default=os.path.join(PROJECT_DIR, "skus.csv"),
                        help="input CSV (default: skus.csv)")
    parser.add_argument("--asin", default=None,
                        help="process ONLY this ASIN from the input file (targeting for a "
                             "single-ASIN proof run). Rows for other ASINs are excluded from "
                             "the run entirely, not counted as failures.")
    parser.add_argument("--fixtures", default=FIXTURES_DIR,
                        help="folder of local ASIN.html fixtures, parsed offline (default: fixtures/)")
    parser.add_argument("--mock-only", action="store_true",
                        help="skip fixture/cache/live/browser lookup entirely; exact Phase 1 behavior")
    parser.add_argument("--live", action="store_true",
                        help="fetch missing (uncached) ASINs directly from Amazon AU (raw GET) - no "
                             "proxy, no Decodo, no concurrency, tiny pilot only")
    parser.add_argument("--limit", type=int, default=None,
                        help="max number of live fetches this run (default: 5 in --live mode)")
    parser.add_argument("--delay-seconds", type=float, default=5.0,
                        help="seconds to sleep between live fetches (default: 5)")
    parser.add_argument("--timeout-seconds", type=float, default=20.0,
                        help="live fetch timeout in seconds (default: 20)")
    parser.add_argument("--cache-dir", default=CACHE_DIR,
                        help="folder for cached raw-fetched HTML (default: page_cache/)")
    parser.add_argument("--debug-cache", action="store_true",
                        help="DEV ONLY: keep every fetched page under --cache-dir, and allow "
                             "reading them back. Off by default: production discards HTML after "
                             "parsing, because keeping it is huge on disk and a cached page "
                             "reports a stale ETA.")
    parser.add_argument("--use-cache", action="store_true",
                        help="allow reusing previously saved HTML instead of fetching. Off by "
                             "default so a normal run always gets fresh ETAs.")
    parser.add_argument("--failure-cache-max", type=int, default=50,
                        help="max anomalous pages (parse errors, blocked/Server Busy, unknown "
                             "layouts) kept under debug_failures/ for inspection; oldest is "
                             "rotated out. 0 disables (default: 50)")
    parser.add_argument("--resume", metavar="RUN_DIR", default=None,
                        help="skip ASINs already completed in a previous run folder, using its "
                             "progress.jsonl. Lets a large batch continue after a crash without "
                             "re-fetching what already succeeded.")
    parser.add_argument("--proxy", default=os.environ.get("SHIPPING_PROXY_URL"),
                        help="route all requests through this proxy, e.g. "
                             "http://user:pass@gate.decodo.com:7000 . Use an AU residential "
                             "exit: it lifts Amazon's Server Busy throttling AND puts the "
                             "delivery context in Australia. Defaults to $SHIPPING_PROXY_URL "
                             "so credentials never need to appear on the command line or in "
                             "this repo.")
    parser.add_argument("--max-block-rate", type=float, default=0.30,
                        help="abort the run if this fraction of fetches come back blocked "
                             "(default: 0.30)")
    parser.add_argument("--block-rate-min-samples", type=int, default=10,
                        help="don't judge the block rate until this many fetches have run "
                             "(default: 10)")
    parser.add_argument("--max-proxy-auth-failures", type=int, default=3,
                        help="abort after this many 407s from the proxy (default: 3)")
    parser.add_argument("--au-postcode", default="2000",
                        help="Australian delivery postcode to establish once per run, so ETA and "
                             "delivery fee are computed for Australia even when the tool runs "
                             "offshore (default: 2000 = Sydney NSW). Empty string disables.")
    parser.add_argument("--browser-live", action="store_true",
                        help="fetch missing (uncached) ASINs via a local headless browser (JS executed) "
                             "- no proxy, no concurrency, tiny pilot only")
    parser.add_argument("--browser-limit", type=int, default=1,
                        help="max number of browser fetches this run (default: 1)")
    parser.add_argument("--browser-delay-seconds", type=float, default=5.0,
                        help="seconds to sleep between browser fetches (default: 5)")
    parser.add_argument("--browser-timeout-seconds", type=float, default=30.0,
                        help="browser fetch timeout in seconds (default: 30)")
    parser.add_argument("--rendered-cache-dir", default=RENDERED_CACHE_DIR,
                        help="folder for cached browser-rendered HTML (default: rendered_cache/)")
    args = parser.parse_args(argv)

    modes_set = sum([args.mock_only, args.live, args.browser_live])
    if modes_set > 1:
        raise SystemExit("--mock-only, --live, and --browser-live are mutually exclusive")

    fixtures_dir = None if args.mock_only else Path(args.fixtures)
    cache_dir = None if args.mock_only else Path(args.cache_dir)
    rendered_cache_dir = None if args.mock_only else Path(args.rendered_cache_dir)

    live_cfg = None
    if args.live:
        live_cfg = {
            "limit": args.limit if args.limit is not None else 5,
            "delay_seconds": args.delay_seconds,
            "timeout_seconds": args.timeout_seconds,
            "max_block_rate": args.max_block_rate,
            "block_rate_min_samples": args.block_rate_min_samples,
            "max_proxy_auth_failures": args.max_proxy_auth_failures,
        }

    browser_cfg = None
    if args.browser_live:
        chrome_binary = find_chrome_binary()
        if chrome_binary is None:
            raise SystemExit(
                "Browser mode requires a local browser. No Chrome/Chromium binary found at "
                "any of the usual locations. Install one of:\n"
                "  - Google Chrome (https://www.google.com/chrome/), or\n"
                "  - Chromium, or\n"
                "  - pip install playwright && playwright install chromium\n"
                "Nothing was installed automatically. Re-run with --browser-live once one is "
                "available, or point PATH at your existing Chrome/Chromium binary.")
        browser_cfg = {
            "limit": args.browser_limit,
            "delay_seconds": args.browser_delay_seconds,
            "timeout_seconds": args.browser_timeout_seconds,
            "chrome_binary": chrome_binary,
        }

    # fixtures/ is only consulted when neither --live nor --browser-live is set
    # (see classify()), so its absence only matters in that (default) mode.
    if fixtures_dir is not None and live_cfg is None and browser_cfg is None and not fixtures_dir.is_dir():
        raise SystemExit("Fixtures folder not found: %s (use --mock-only to skip fixture lookup)"
                         % fixtures_dir)

    started = dt.datetime.now()
    # Suffix on collision: two runs in the same minute must not overwrite each
    # other's audit files.
    run_id = started.strftime("%Y-%m-%d_%H%M")
    run_dir = os.path.join(RUNS_DIR, run_id)
    suffix = 2
    while os.path.exists(run_dir):
        run_id = "%s_%d" % (started.strftime("%Y-%m-%d_%H%M"), suffix)
        run_dir = os.path.join(RUNS_DIR, run_id)
        suffix += 1
    os.makedirs(run_dir)
    os.makedirs(RESULTS_DIR, exist_ok=True)

    log_path = os.path.join(run_dir, "run.log")
    log_lines = []

    def log(message):
        line = "%s  %s" % (dt.datetime.now().isoformat(timespec="seconds"), message)
        log_lines.append(line)
        print(line)

    log("run folder: %s" % run_dir)
    log("input: %s" % args.input)
    log("page cache: %s" % (cache_dir if cache_dir is not None else "(disabled, --mock-only)"))

    # Must happen before ANY request, so the AU location cookie and the product
    # fetches all leave from the same exit IP.
    if args.proxy:
        configure_proxy(args.proxy, log)
    elif live_cfg is not None:
        log("proxy: NONE - direct connection. If Amazon returns 'Server Busy' stubs, set "
            "SHIPPING_PROXY_URL to an AU residential proxy.")
    log("rendered cache: %s" % (rendered_cache_dir if rendered_cache_dir is not None else "(disabled, --mock-only)"))
    if browser_cfg is not None:
        log("BROWSER MODE - headless-Chrome direct fetch enabled (%s). No proxy, no concurrency. "
            "limit=%d delay=%.1fs timeout=%.1fs"
            % (browser_cfg["chrome_binary"], browser_cfg["limit"],
               browser_cfg["delay_seconds"], browser_cfg["timeout_seconds"]))
        log("fixtures: (not consulted in --browser-live mode)")
    elif live_cfg is not None:
        log("LIVE MODE - direct fetch enabled. No proxy, no Decodo, no concurrency. "
            "limit=%d delay=%.1fs timeout=%.1fs"
            % (live_cfg["limit"], live_cfg["delay_seconds"], live_cfg["timeout_seconds"]))
        log("fixtures: (not consulted in --live mode)")
    else:
        log("offline mode - no network requests will be made")
        log("fixtures: %s" % (fixtures_dir if fixtures_dir is not None else "(disabled, --mock-only)"))

    if not os.path.exists(args.input):
        raise SystemExit("Input file not found: %s" % args.input)

    rows, stats = load_input(args.input, log)

    if args.asin:
        target_asin = normalise_asin(args.asin)
        before = len(rows)
        rows = [r for r in rows if r["asin"] == target_asin]
        if not rows:
            raise SystemExit("--asin %s not found in %s" % (target_asin, args.input))
        # Excluded rows are out of scope for this run, not failures - so
        # input_count below counts only the targeted rows and reconciliation
        # still balances against exactly what was processed.
        log("--asin %s: processing %d of %d input rows (others excluded from this run)"
            % (target_asin, len(rows), before))

    if args.resume:
        done = load_completed_asins(args.resume, log)
        before = len(rows)
        rows = [r for r in rows if r["asin"] not in done]
        log("--resume: %d of %d rows remain to process" % (len(rows), before))
        if not rows:
            log("--resume: nothing left to do")

    input_count = len(rows)
    rejected = [r for r in rows if r["reject_reason"]]
    units, duplicates_collapsed = dedupe(rows, log)

    # Establish the Australian delivery context ONCE, before any ASIN is fetched,
    # then reuse the session cookie for the whole batch. 10,000 ASINs = 2 setup
    # requests + 10,000 product requests, not 20,000.
    au_context = {"requested": False, "established": False, "setup_bytes": 0,
                  "postcode": args.au_postcode or None}
    # Skip setup entirely when there is nothing to fetch (e.g. a fully-resumed
    # run) - no work means no reason to spend the 2 setup requests.
    if live_cfg is not None and args.au_postcode and units:
        au_context["requested"] = True
        ok, setup_bytes = establish_au_delivery_context(
            args.au_postcode, live_cfg["timeout_seconds"], log)
        au_context["established"] = ok
        au_context["setup_bytes"] = setup_bytes

    retention = {
        "debug_cache": args.debug_cache,
        "read_cache": args.debug_cache or args.use_cache,
        "failure_dir": Path(run_dir) / "debug_failures",
        "failure_max": max(0, args.failure_cache_max),
        "failures_saved": 0,
    }
    if args.debug_cache:
        log("HTML retention: DEBUG-CACHE ON - every page kept under %s (dev only)" % cache_dir)
    else:
        log("HTML retention: production - HTML parsed in memory and discarded; only up to %d "
            "anomalous pages kept for debugging" % retention["failure_max"])
    if retention["read_cache"]:
        log("HTML reuse: ENABLED - previously saved pages may be reparsed (ETAs may be stale)")

    extracted_at = started.replace(microsecond=0).isoformat()
    live_state = {"fetches_done": 0, "wire_bytes": 0, "blocked": 0,
                  "proxy_auth_failures": 0, "aborted": None}
    browser_state = {"fetches_done": 0, "served_hosts": {}}
    progress_path = os.path.join(run_dir, "progress.jsonl")
    records = build_records(units, rejected, extracted_at, fixtures_dir, cache_dir, live_cfg, live_state,
                            rendered_cache_dir, browser_cfg, browser_state, log,
                            retention, progress_path)

    # Route every record to exactly one sheet.
    by_sheet = {name: [] for name in DATA_SHEETS}
    for rec in records:
        sheet = SHEET_FOR_STATUS.get(rec["status"])
        if sheet is None:
            raise SystemExit("Unrouted status %r on input row %s - refusing to drop it silently."
                             % (rec["status"], rec["asin"]))
        by_sheet[sheet].append(rec)

    counts = {name: len(by_sheet[name]) for name in DATA_SHEETS}
    routed_total = sum(counts.values())
    reconciles = routed_total == input_count
    log("counts: " + ", ".join("%s=%d" % (k, v) for k, v in counts.items()))
    log("reconciliation: %d routed vs %d input rows -> %s"
        % (routed_total, input_count, "PASS" if reconciles else "FAIL"))

    # input_normalized.csv - the audit trail for "why wasn't my SKU in the output?"
    normalized_path = os.path.join(run_dir, "input_normalized.csv")
    with open(normalized_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["input_row_number", "sku", "asin_normalized", "rc_listing_id",
                         "title", "dedupe_key", "action"])
        seen = set()
        for row in rows:
            if row["reject_reason"]:
                action = "rejected:%s" % row["reject_reason"]
                key = ""
            else:
                key = row["asin"] or row["sku"]
                if key in seen:
                    action = "duplicate_of_fetch_key"
                else:
                    seen.add(key)
                    action = "fetch"
            writer.writerow([row["input_row_number"], row["sku"], row["asin"],
                             row["rc_listing_id"], row["title"], key, action])
    log("wrote %s" % normalized_path)

    if browser_cfg is not None:
        phase_label = "4 - tiny browser-rendered pilot"
    elif live_cfg is not None:
        phase_label = "3 - tiny live pilot"
    else:
        phase_label = "2 - offline fixtures + mock fallback"

    summary = OrderedDict([
        ("run_id", run_id),
        ("phase", phase_label),
        ("started_at", extracted_at),
        ("input_file", os.path.basename(args.input)),
        ("fixtures_dir", str(fixtures_dir)
         if (fixtures_dir is not None and live_cfg is None and browser_cfg is None) else None),
        ("cache_dir", str(cache_dir) if cache_dir is not None else None),
        ("rendered_cache_dir", str(rendered_cache_dir) if rendered_cache_dir is not None else None),
        # The INTENDED basis. What was actually observed per row is in each row's
        # delivery_location_basis column and can differ - Amazon geolocates a
        # logged-out session from the exit IP. Never read this as confirmation.
        ("delivery_location_basis_configured", DELIVERY_LOCATION_BASIS),
        ("delivery_location_basis_observed",
         sorted({rec["delivery_location"] for rec in records if rec["delivery_location"]}) or None),
        ("input_count", input_count),
        ("blank_rows_skipped", stats["blank_rows_skipped"]),
        ("invalid_rows", len(rejected)),
        ("unique_fetch_units", len(units)),
        ("duplicate_fetches_avoided", duplicates_collapsed),
        ("counts", counts),
        ("reconciliation", "PASS" if reconciles else "FAIL"),
        ("reconciliation_detail",
         "%d routed == %d input rows" % (routed_total, input_count)),
        ("live_mode", live_cfg is not None),
        ("live_fetch_limit", live_cfg["limit"] if live_cfg is not None else None),
        ("html_debug_cache", args.debug_cache),
        ("html_cache_reads_allowed", retention["read_cache"]),
        ("html_discarded_after_parse", not args.debug_cache),
        ("failure_html_cap", retention["failure_max"]),
        ("failure_html_saved", retention["failures_saved"]),
        ("resumed_from", args.resume),
        ("au_postcode_requested", au_context["postcode"]),
        ("au_context_established", au_context["established"]),
        ("au_context_setup_requests", 2 if au_context["requested"] else 0),
        ("au_context_setup_bytes", au_context["setup_bytes"]),
        ("live_fetches_done", live_state["fetches_done"]),
        ("live_wire_bytes_total", live_state["wire_bytes"]),
        ("live_wire_kb_per_fetch",
         round(live_state["wire_bytes"] / 1024.0 / live_state["fetches_done"], 1)
         if live_state["fetches_done"] else None),
        ("browser_mode", browser_cfg is not None),
        ("browser_fetch_limit", browser_cfg["limit"] if browser_cfg is not None else None),
        ("browser_fetches_done", browser_state["fetches_done"]),
        ("browser_served_hosts", browser_state["served_hosts"] or None),
        ("target_asin", normalise_asin(args.asin) if args.asin else None),
        ("browser_binary", browser_cfg["chrome_binary"] if browser_cfg is not None else None),
        ("network_calls", live_state["fetches_done"] + browser_state["fetches_done"]),
        ("proxy_used", PROXY_STATE["configured"]),
        ("proxy_endpoint", PROXY_STATE["host"]),
        ("blocked_by_amazon", live_state["blocked"]),
        ("proxy_auth_failures", live_state["proxy_auth_failures"]),
        ("block_rate", round(live_state["blocked"] / live_state["fetches_done"], 3)
         if live_state["fetches_done"] else None),
        ("aborted_reason", live_state["aborted"]),
        ("decodo_used", False),
        ("db_writes", 0),
    ])

    filename = "shipping_enrichment_%s_%d-input_%d-extracted.xlsx" % (
        started.strftime("%Y-%m-%d"), input_count, counts["extracted"])
    xlsx_path = unique_path(os.path.join(RESULTS_DIR, filename))

    sheets = [("summary", build_summary_rows(summary))]
    for name in DATA_SHEETS:
        rows_out = [COLUMNS] + [[rec[c] for c in COLUMNS] for rec in by_sheet[name]]
        sheets.append((name, rows_out))

    writer_used = write_xlsx(xlsx_path, sheets, log)
    summary["excel_writer"] = writer_used
    summary["output_file"] = os.path.basename(xlsx_path)
    log("wrote %s" % xlsx_path)

    # CSV review files - the Excel workbook is hard to open/check quickly, so every
    # sheet also lands as a plain CSV in the run folder, built from the exact same
    # rows (sheets / records), never re-derived. summary.json remains the machine-
    # readable summary; summary.csv is the same summary metrics for quick viewing.
    csv_filenames = {
        "summary": "summary.csv",
        "extracted": "extracted.csv",
        "redirects": "redirects.csv",
        "no_required_data": "no_required_data.csv",
        "errors": "errors.csv",
    }
    for sheet_name, rows_out in sheets:
        header, data_rows = rows_out[0], rows_out[1:]
        write_csv_rows(os.path.join(run_dir, csv_filenames[sheet_name]), header, data_rows, log)

    # review_all_rows.csv - every non-summary row in one file, in original input
    # order, with a leading `sheet` column so nothing needs to be cross-referenced
    # across four separate files to see the whole run at a glance.
    review_header = ["sheet"] + COLUMNS
    review_rows = [[SHEET_FOR_STATUS[rec["status"]]] + [rec[c] for c in COLUMNS] for rec in records]
    write_csv_rows(os.path.join(run_dir, "review_all_rows.csv"), review_header, review_rows, log)

    summary_path = os.path.join(run_dir, "summary.json")
    with open(summary_path, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)
    log("wrote %s" % summary_path)

    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(log_lines) + "\n")

    if not reconciles:
        print("RECONCILIATION FAILED - see the summary sheet.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
