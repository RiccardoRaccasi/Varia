# nav-tracker

Tracks the Net Asset Value of the **Stableton Morningstar PitchBook Unicorn 20 AMC
All Investors** certificate (ISIN `CH1234846777`, USD) in an Excel workbook, and
appends the current price from [Swiss Fund Data][src] once a day.

[src]: https://www.swissfunddata.ch/sfdpub/en/funds/show/175908

```
nav-tracker/
├── nav_tracker.py          the whole thing — library + CLI, stdlib + openpyxl
├── run_daily.sh            what the scheduler calls
├── requirements.txt
├── data/
│   ├── unicorn20_nav_history.xlsx      the tracking file
│   └── swissfunddata_export_*.csv      the history it was seeded from
├── schedule/               cron · launchd · systemd + an installer
└── tests/                  offline tests, no network needed
```

## Before the first scheduled run: confirm the page

Everything below is verified against saved fixtures, **but the live page's exact
markup has not been seen** — it is blocked from the network this was built on.
One command tells you whether the scraper reads the right number:

```bash
python3 nav_tracker.py probe
```

It prints every price-looking figure on the page, best first, and the top one
marked `>>` is what `update` will record. If that is the Net Asset Value with
the right date, you are done — nothing to change. If a different row should be
used, pin it once:

```bash
python3 nav_tracker.py probe --label-regex 'redemption price'   # try it
./run_daily.sh --label-regex 'redemption price'                 # then use it
```

If `probe` finds nothing at all, save the page and look at it:

```bash
python3 nav_tracker.py probe --save-html /tmp/fund.html
```

## Everyday use

```bash
python3 nav_tracker.py update          # fetch the price, append it   <- the routine
python3 nav_tracker.py fetch           # fetch only, write nothing
python3 nav_tracker.py show -n 10      # tail the stored series
python3 nav_tracker.py update --dry-run --json
```

Re-running is safe. A date already in the file is left alone; if the fund
restates a value, the existing row is corrected rather than duplicated; rows
stay sorted by date whatever order they arrive in.

## Schedule it for 10:00

**On a machine that is on at 10:00** (macOS launchd, Linux cron):

```bash
./schedule/install.sh              # ./schedule/install.sh uninstall  to remove
```

Both fire at 10:00 in the machine's local timezone — set it to `Europe/Zurich`
for 10:00 Swiss time. `schedule/` also holds a systemd timer and a hand-editable
crontab line. Logs land in `logs/nav_tracker.log`.

**Without a machine to keep on**, use `.github/workflows/nav-daily.yml`: it runs
in GitHub Actions at 10:00 Europe/Zurich (handling the CET/CEST shift on its own)
and commits each new observation. It starts working once it is on the repo's
default branch; until then it is inert.

## Calling it from other code

The module is importable and has no global state:

```python
import sys; sys.path.insert(0, "/path/to/Varia/nav-tracker")
import nav_tracker

point = nav_tracker.fetch_current_price()     # PricePoint(date, value, currency, ...)
print(point.date, point.value, point.currency)

result = nav_tracker.update()                 # scrape + append
if result.changed:
    print(result.message)                     # "recorded 2026-09-10 = 232.5000 USD"

for row in nav_tracker.read_history():        # the full series, oldest first
    ...
```

Everything takes explicit arguments, so a second fund needs no edits to the file:

```python
nav_tracker.update(url="https://www.swissfunddata.ch/sfdpub/en/funds/show/123456",
                   workbook="/data/other_fund.xlsx")
```

For non-Python callers, every command speaks JSON on stdout with `--json`, and
the exit code carries the outcome:

| code | meaning |
|-----:|---------|
| 0 | fine — appended, corrected, or already up to date |
| 2 | the page could not be fetched or no price could be read |
| 3 | a price was read but it failed the plausibility check |

`NAV_TRACKER_URL` and `NAV_TRACKER_WORKBOOK` override the defaults via the
environment, which is usually the tidiest way to point the scheduled job at a
workbook kept outside the repo.

## The workbook

Sheet **NAV History** — one row per publication, oldest first, plus a line chart:

| Date | Net Asset Value | Currency | Source | Retrieved (UTC) |
|------|----------------:|----------|--------|-----------------|
| 08/09/2026 | 231.856 | USD | csv-import | 2026-09-14T09:32:44+00:00 |
| 09/09/2026 | 231.845 | USD | swissfunddata | 2026-09-14T10:00:03+00:00 |

`Source` says where each row came from, so a scraped value is always
distinguishable from the seeded history. Sheet **Fund** carries the fund name,
ISIN, currency, source URL and the range covered.

Rebuilding from scratch, or topping up from a fresh CSV export:

```bash
python3 nav_tracker.py import-csv data/swissfunddata_export_2026-09-09.csv
python3 nav_tracker.py import-csv newer_export.csv --merge   # keep scraped rows
```

The workbook is rewritten atomically (temp file, then rename), so an interrupted
run cannot leave a half-written file.

## What stops a bad scrape

A mis-read number would quietly poison the series, so the routine refuses rather
than guesses:

- **Labels are matched, not positions** — a layout change breaks loudly instead
  of silently reading the neighbouring column.
- Fund volume, units outstanding, performance figures, ISINs and valor numbers
  are never treated as prices.
- Swiss, German, French and English number formats are all understood
  (`1'234.56`, `1.234,56`, `1 234,56`). Text glued together by markup is
  **refused**, not guessed at.
- A price with **no date on the page is refused** — it cannot be filed against a
  day. `--assume-today` overrides this, and then an unchanged value is treated as
  a stale page rather than appended.
- A move of more than **30%** from the last stored value is rejected
  (`--max-move` to change, `--force` to override).
- A future-dated price is rejected.

## Tests

```bash
python3 tests/test_nav_tracker.py
```

32 tests, no network: page-parsing across six layouts, number and date parsing,
and the workbook round-trip including idempotency, restatement, out-of-order
dates and every refusal above.
