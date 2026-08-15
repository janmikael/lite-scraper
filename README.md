# Shipping Enrichment Lite — how to run

Standalone Amazon AU tool: **ASIN → delivery ETA + merchant + delivery fee.**
No database, no scheduler, no browser. One HTTP request per ASIN.
Needs an AU residential proxy (see §2b) — Amazon blocks direct connections.

## 1. Put your ASINs in `skus.csv`

Open `skus.csv` and paste one ASIN per line under the header:

```
sku,asin,rc_listing_id,title
B0CP4XY9QC
B0D54PM77N
B07Y51PCW7
```

A bare ASIN per line is fine — it's used as both SKU and ASIN. If you have real
SKUs, use the full form instead: `MY-SKU-1,B0CP4XY9QC,RC-1001,Some product`.
Only `sku` and `asin` matter; `rc_listing_id` and `title` are optional pass-through.

## 2. Run it

```bash
cd /path/to/lite-scraper
python3 run_shipping_enrichment.py --live --limit 10 --delay-seconds 5
```

Requires **Python 3.7+** and nothing else — no `pip install`, no virtualenv. Every
import is standard library. (`openpyxl` is used for the Excel file if you happen to
have it; if not, a built-in fallback writer produces the same workbook.)

`--live` is required to actually fetch. **Without it the tool runs offline and
fetches nothing** — that default is deliberate, so you can never spend bandwidth
by accident.

| Flag | Default | What it does |
|---|---|---|
| `--live` | off | Actually fetch from Amazon AU. Required. |
| `--limit N` | 5 | **Max ASINs fetched this run.** Your main safety valve. |
| `--delay-seconds N` | 5 | Pause between requests. Don't lower it much. |
| `--input FILE` | `skus.csv` | Use a different input file. |
| `--au-postcode` | `2000` | AU delivery postcode (Sydney). Set once per run. |
| `--asin B0XXXXXXXX` | — | Process only this one ASIN. Handy for spot-checks. |
| `--resume RUN_DIR` | — | Skip ASINs already done in an earlier run folder. |
| `--failure-cache-max N` | 50 | Cap on saved failure pages. `0` disables. |
| `--debug-cache` | off | DEV ONLY: keep every page's HTML. Huge on big batches. |
| `--use-cache` | off | Reuse previously saved HTML instead of fetching (stale ETAs). |
| `--proxy URL` | `$SHIPPING_PROXY_URL` | Route via proxy. Needed — see §2b. |
| `--max-block-rate N` | 0.30 | Abort if this share of fetches are blocked. |
| `--max-proxy-auth-failures N` | 3 | Abort after this many proxy 407s. |

Start small (`--limit 10`), confirm the output looks right, then raise it.

## 2b. Proxy (needed — Amazon blocks direct connections)

Fetching direct gets Amazon's **"Server Busy"** anti-bot stub (~2.8 KB) instead of the
product page — measured at roughly a 75% block rate. An **Australian residential
proxy** fixes both that and the delivery geography in one go.

Set the credentials as an environment variable so they never touch this repo:

```bash
export SHIPPING_PROXY_URL='http://USER:PASS@gate.decodo.com:7000'
python3 run_shipping_enrichment.py --live --limit 20 --delay-seconds 5
```

Use an **AU exit**. Decodo encodes that in the username (e.g. `user-XXXX-country-au`)
— check your dashboard for the exact syntax. A sticky/session exit is better than
per-request rotation: the delivery-location cookie is tied to the IP that set it.

The run aborts itself if things go wrong, so a bad proxy can't burn the whole list:

| Guard | Default | Flag |
|---|---|---|
| Block rate too high | abort above 30% (after 10 fetches) | `--max-block-rate` |
| Proxy rejects credentials (407) | abort after 3 | `--max-proxy-auth-failures` |

Aborts are not silent: partial results are still written, `summary.json` records
`aborted_reason`, and `--resume` picks up where it stopped. Credentials are scrubbed
from all logs and output.

## 3. Where the results go

Two places, written at the end of every run:

**`results/`** — the Excel file to send on:
```
results/shipping_enrichment_2026-08-13_10-input_9-extracted.xlsx
```
The filename shows the date, how many rows went in, and how many produced data.
Sheets: `summary`, `extracted`, `redirects`, `no_required_data`, `errors`.

**`runs/<date>_<time>/`** — the same data as plain CSVs, easier to open and diff:
```
extracted.csv          rows that worked
redirects.csv          ASIN pointed somewhere else
no_required_data.csv   page loaded, no usable delivery info
errors.csv             fetch/validation failures
review_all_rows.csv    everything in one file  <- usually the one you want
summary.csv            counts + run settings
summary.json           same, machine-readable
run.log                what happened, line by line
```

Progress also prints to the terminal live as it runs.

**Every input row lands in exactly one sheet** and the counts are reconciled — if
they don't balance, the summary says `FAIL`. Nothing is ever silently dropped.

## 4. Output columns

`sku, asin, eta_raw, eta_date, merchant, delivery_fee, checked_at, status, error, delivery_location`

- `delivery_fee` — `0.0` means free delivery. **Blank means unknown, not free.**
- `eta_date` — normalised latest date (`2026-08-17`). Blank if it couldn't be
  parsed confidently; `eta_raw` always keeps Amazon's original wording.
- `delivery_location` — **check this.** It must say `2000`. If it says
  `Philippines` (or anywhere non-AU), the ETA and fee on that row are for the
  wrong country and are not usable. The run log prints a WARNING when this happens.

## 5. HTML handling

By default **no page HTML is kept**. Each page is parsed in memory and discarded,
so a run leaves behind only the CSV/XLSX results plus a small checkpoint. You do
not need to clean anything between runs, and every run fetches fresh ETAs.

Two deliberate exceptions:

- **Failure samples.** Up to 50 genuinely anomalous pages (parser failures,
  blocked/Server Busy, unrecognised layouts) are kept in
  `runs/<timestamp>/debug_failures/` so you can see what went wrong. The oldest is
  deleted once the cap is hit, so this can't grow. Tune with
  `--failure-cache-max N`, or `0` to switch it off. Ordinary outcomes — out of
  stock, redirects, no fee found — are never saved.
- **`--debug-cache`** keeps every page, for development. Don't use it on a big
  batch: at ~2.4 MB decompressed per page, 200k ASINs would be roughly 480 GB.

Cached pages are also **never read back automatically**, because a cached page is
a snapshot and delivery ETAs change daily. Pass `--use-cache` if you deliberately
want to re-parse pages you already saved.

## 6. Resuming a big batch

Every completed ASIN is appended to `runs/<timestamp>/progress.jsonl` as it
finishes. If a long run dies partway, continue it with:

```bash
python3 run_shipping_enrichment.py --live --limit 200000 --resume runs/2026-08-13_1951
```

Already-completed ASINs are skipped and not re-fetched. If there's nothing left to
do, the run exits without even the location setup.

## Cost per run

- One-time AU delivery-location setup: 2 requests, ~110 KB
- Per ASIN: 1 request, ~430 KB transferred (not stored)

10,000 ASINs ≈ 10,002 requests ≈ 4.1 GB downloaded.

**Disk** is far smaller, because HTML isn't kept. For 200,000 ASINs, excluding the
CSV/XLSX results: ~40 MB of `progress.jsonl` + run CSVs, plus at most ~120 MB of
capped failure samples — under ~160 MB total, and flat no matter how large the
batch gets.
