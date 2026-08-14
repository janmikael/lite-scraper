# fixtures/

Local, manually-saved Amazon AU HTML pages. The tool parses these **offline** — it never
fetches a page itself. No network request is ever made from this folder.

## How to add a real fixture later

1. Open the product page in a normal browser: `https://www.amazon.com.au/dp/<ASIN>`.
2. Save the page HTML (e.g. `File > Save Page As...`, "Webpage, HTML only" is fine — the
   parser only needs the raw markup, not saved images/CSS).
3. Name the file by the **requested** ASIN, uppercase, with a `.html` extension:
   `B0TEST1234.html`.
4. Put the file in this folder (`fixtures/`).
5. Add a row for that SKU/ASIN to `skus.csv` (or your real input file) if it isn't already
   there.
6. Run the script offline as usual:
   ```
   python3 run_shipping_enrichment.py
   ```
   For each normalized ASIN, the tool looks for `fixtures/<ASIN>.html`. If the file exists,
   it parses real fields from that saved HTML. If it doesn't, the row falls back to the
   built-in mock data (Phase 1 behavior) or, if there's no mock entry either, is routed to
   `no_required_data` with a clear reason — never silently dropped.

There is no network fallback anywhere in this path. A missing fixture never triggers a
fetch.

## What the parser looks for

Conservative, standard-library-only, regex-based extraction (see `run_shipping_enrichment.py`
for the exact patterns). If a field can't be found confidently, it is left **blank** — the
parser never guesses or invents a value:

- **Canonical/final ASIN** — a `<link rel="canonical" href=".../dp/ASIN">` tag. If it names a
  different ASIN than the one requested, the row is routed to `redirects` and no further
  fields are extracted from that page (see PLAN.md §9 — redirects are recorded, not
  auto-followed, on purpose).
- **Sold by** — an element with `id="sellerProfileTriggerId"`, or the text of
  `id="merchantInfoFeature_feature_div"` with a leading "Sold by" stripped.
- **Ships from** — the text of `id="fulfillerInfoFeature_feature_div"` with a leading
  "Ships from" stripped.
- **Delivery message** — the text of `id="deliveryBlockMessage"` (or
  `id="mir-layout-DELIVERY_BLOCK"` as a fallback).
- **Delivery fee** — parsed out of the delivery message text only when unambiguous:
  the word "free" (case-insensitive) anywhere in the message becomes `0.0`; a clear
  `$12.34`-style figure becomes that number. Anything else is left **blank**, not `0`.
- **ETA date** — a `"<Weekday>, <day> <Month>"` pattern in the delivery message (Amazon AU's
  usual format, e.g. "Wednesday, 20 May"). The day count to that date, and whether it's
  over 14 days out, are computed from today's date. No year is on the page, so the nearest
  future occurrence of that month/day is assumed.
- **Availability** — the text of `id="availability"`, mapped to `IN_STOCK` /
  `OUT_OF_STOCK` / `LIMITED` on a small set of known phrases.

These are the same page regions the real Amazon AU product page uses for this data. They are
**not guaranteed to match every page layout** — Amazon changes markup over time and tests
variants. Treat every field this parser produces as "best effort from what this page
happened to contain," and expect to extend the patterns once real saved pages are in this
folder. That is expected iteration, not a bug.

## The three sample fixtures

`sample_extracted.html`, `sample_redirect.html`, and `sample_no_required_data.html` are
**small synthetic files I wrote by hand**, not real Amazon pages, and not proof that
extraction works against Amazon's actual markup. They exist purely so the parsing code has
something deterministic to run against — one file per output sheet (`extracted`,
`redirects`, `no_required_data`).

The tool only ever looks up `fixtures/<ASIN>.html` (see "How to add a real fixture later"
above) — it has no notion of a `sample_*.html` filename. So each sample also exists as an
identical **ASIN-named copy** that `skus.csv` actually points at:

| Descriptive name (for reading) | ASIN-named copy (what the tool loads) | Referenced by `skus.csv` row |
|---|---|---|
| `sample_extracted.html` | `B0FIX00001.html` | `SKU-0006` |
| `sample_redirect.html` | `B0FIX00002.html` | `SKU-0007` |
| `sample_no_required_data.html` | `B0FIX00003.html` | `SKU-0008` |

`SKU-0009` (`B0FIX00004`) deliberately has **no** fixture file and no mock entry, to prove
the missing-fixture path routes to `no_required_data` with a clear reason instead of
vanishing. If you edit a `sample_*.html` file, copy your changes into its `B0FIXxxxxx.html`
counterpart too — they're kept as plain duplicate files, not a symlink or a build step, so
there is nothing here to run before the script sees your edit; forgetting the copy is a
correctness bug, but no more.

See `plans/2026-08-05-phase-2-fixture-extraction-plan.md` for what each one is meant to
prove.

`sample_extracted.html`'s delivery date is a real weekday/date pair computed relative to
2026-08-05 (7 days out, so it lands comfortably under the 14-day threshold at authoring
time). If you're reading this long after that and the automated self-check starts failing
because the fixture's date has drifted into "over 14 days" or into the past, that's expected
staleness in a hand-picked calendar date — bump the date string inside the file, it isn't a
parser bug.
