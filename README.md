# OTIF Auto

Builds Nutaste's **OTIF** (On Time In Full) report from Business Central.

It pulls sales orders, shipments and order headers from the BC OData service,
computes OTIF line by line, and writes Parquet files for Power BI plus an Excel
workbook. A full run takes about 90 seconds.

It replaces a Power Query workbook that recalculated 1.16 million formula cells
on every refresh and eventually stopped opening. Those formulas are now about 15
vectorised pandas operations, and every lookup is a hash join rather than a
per-row scan.

---

## Contents

- [Quick start](#quick-start)
- [What OTIF means here](#what-otif-means-here)
- [How a run works](#how-a-run-works)
- [Outputs](#outputs)
- [Configuration](#configuration)
- [The two reports](#the-two-reports)
- [Constraints and gotchas](#constraints-and-gotchas)
- [Scheduling](#scheduling)
- [Power BI](#power-bi)
- [Knowing when it last refreshed](#knowing-when-it-last-refreshed)
- [Troubleshooting](#troubleshooting)
- [Repository layout](#repository-layout)

---

## Quick start

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
copy .env.example .env
```

Edit `.env` and set at minimum:

- `BC_USERNAME` — your Business Central user name.
- `BC_ACCESS_KEY` — the **Web Service Access Key** from the BC user card
  (User Settings → Users → your user → Web Service Access Key), *not* your
  Windows password. Generate one if the field is empty.
- `PARQUET_DIR` — where Power BI reads from; see [Power BI](#power-bi).

Check the business logic without touching the ERP — it runs on fixtures, needs
no credentials, and takes a second:

```bash
.venv\Scripts\python selftest.py
```

Check the connection:

```bash
.venv\Scripts\python probe.py
```

Then run it:

```bash
.venv\Scripts\python main.py
```

---

## What OTIF means here

A sales order line is **OTIF** if the quantity delivered *on or before its
deadline* is at least the quantity ordered.

```
Internal_Deadline        = PO_Date + lead days        (7 if MTA, else 14)
Requested_Shipment_Date  = Requested_Delivery_Date - 4 days
                           (falls back to PO_Date when blank)

Final_Deadline           = MAX(Internal_Deadline, Requested_Shipment_Date)

Qty_Delivered_Before_Deadline
                         = SUM(shipment qty WHERE Posting_Date <= Final_Deadline)

OTIF                     = Qty_Delivered_Before_Deadline >= Quantity
```

Every constant above lives in `config.py` (`LEAD_DAYS_MTA`,
`LEAD_DAYS_NON_MTA`, `SHIPMENT_LEAD_OFFSET_DAYS`). Whether an item is MTA comes
from `reference/mta_list.csv`.

Because the quantity sum is bounded by the deadline, **timeliness is already
baked into the OTIF flag** — a delivery posted after the deadline contributes
nothing. This is also why a line's OTIF result is final the moment its deadline
passes.

### Reporting grain

OTIF is reported by **deadline month**: a line whose `Final_Deadline` falls in
July is a July line, regardless of when it was ordered or delivered.

### OTIF is only assessable once the deadline passes

A line that is undelivered but **not yet due** is pending, not failed. Scoring it
as a failure understates OTIF and makes the current month look artificially bad.
So the pipeline publishes three views:

| Table | Contents | Use it for |
|---|---|---|
| `otif_lines` | every line, future deadlines included | the complete operational picture |
| `snop` | one row per order+item, **deadline already elapsed** | **month-by-month OTIF %** |
| `pending` | only lines whose deadline has *not* arrived | the forward order book / expediting |

`snop` is the one your OTIF percentages should come from. The cutoff is the run
date, controlled by `SNOP_REQUIRE_ELAPSED_DEADLINE` (leave it true).

Three columns make the status explicit so no downstream date arithmetic is
needed:

| Column | Meaning |
|---|---|
| `Deadline_Status` | `Elapsed` / `Pending` / `No deadline` |
| `Days_To_Deadline` | signed days to `Final_Deadline`; negative means past |
| `Qty_Short` | `Quantity - Qty_Delivered_Before_Deadline`, floored at 0 |

`No deadline` means neither a PO date nor a requested delivery date exists, so
nothing can be derived. These count as failures — you cannot demonstrate on-time
delivery against no commitment date — but they are really a data-entry problem in
BC worth fixing at source. It is a few dozen lines.

### `On_Time` is not an independent measure

`In_Full`, `On_Time` and `OTIF` are currently all the same value. Timeliness is
enforced inside the quantity sum, so the OTIF figure is sound, but `On_Time`
tells you nothing `In_Full` does not. If you want a genuine three-way split the
natural definitions are:

- `On_Time` = delivered *anything* by `Final_Deadline`
- `In_Full` = total delivered (any date) >= `Quantity`
- `OTIF` = both

That is a three-line change in `transform.build_otif_lines`.

---

## How a run works

```
extract.py     BC OData  ->  typed DataFrames
                 archive order lines   (the spine)
                 open order lines      (live, not yet archived)
                 shipment lines + headers
                 order headers         (live + archive)
                 reference CSVs        (MTA flags, pincodes)
                        |
transform.py   collapse duplicates -> build keys -> resolve PO date
               -> deadlines -> delivered-before-deadline -> OTIF flags
               -> split into otif_lines / snop / pending
                        |
outputs.py     Parquet (current/ + dated history/) and values-only Excel
```

`main.py` orchestrates it, logs a summary, and writes a run record.

### Where the data comes from

| Source entity | Gives us |
|---|---|
| `ArchiveSalesOrdeLinesBi` | order lines for **closed** orders — the bulk of the spine |
| `SaleOrdersLineBi` | order lines for **open** orders |
| `PostedSalesShipmentLinesBi` | delivered quantities |
| `PostedSalesShipmentBi` | shipment posting dates + the order they belong to |
| `SalesOrderArchivesHeaderBi` | PO date, brand, vertical, pincode (closed orders) |
| `SalesOrder` | the same header fields for open orders |
| `reference/mta_list.csv` | MTA flag per item (6,567 items, hand-maintained) |
| `reference/pincode_mapping.csv` | pincode → city / state (277 rows, hand-maintained) |

The two CSVs existed only inside the old workbook and have no source in BC. Keep
editing them by hand.

### Row identity

Two composite keys, both joined with `|` (`transform.KEY_SEP`):

- `Key` = `Document_No | Line_No | No | Quantity` — the de-duplication key.
- `Key_Req` = `Document_No | No` — the grain shipments join on, and the grain of
  `snop`.

**The delimiter matters.** Concatenated without one, document `SO1` line `1` and
document `SO` line `11` produce the same key, and because `Key` drives
de-duplication that silently deletes a real order line. `selftest.py` covers it.

A line can arrive from both sources — live while its order is open, archived once
it closes. Those are collapsed to one row, archive preferred, keyed on
**document + line + item**. `Key` cannot do that job because it includes the
quantity, so a line edited before its order closed yields two different keys and
would be counted twice. The item number has to be part of the identity because
**BC reuses line numbers within a document** — `SO-M12803` line 10000 carries two
different items, and keying on document+line alone deletes a real line.

---

## Outputs

Written to `PARQUET_DIR`:

```
current/          latest run, overwritten each time
  PO_otif_lines.parquet      every order line
  PO_snop.parquet            one row per order+item, elapsed deadlines only
  PO_pending.parquet         deadline not yet reached
  PO_sales_orders.parquet    order header dimension (brand, vertical, city)
  PO_shipment_lines.parquet  delivered quantities
  release_*.parquet          the second report, see below

history/          one dated snapshot per table per run (KEEP_HISTORY)
  PO_otif_lines/PO_otif_lines_2026-08-26.parquet
  ...

last_run.json     status of the most recent run, success or failure
run_history.jsonl one line appended per run
```

Every table also carries `Snapshot_Date` (the reporting day, for slicing
history) and `Refreshed_At` (the moment the data was pulled).

Excel goes to `OUTPUT_DIR` (default `./output`) as `PO OTIF Report.xlsx` — values
only, three sheets: OTIF Lines, SNOP, Pending.

The `PO` prefix (`config.EXCEL_NAME`, `config.PARQUET_PREFIX`) marks this
report's files apart from the release-date report's, which share the same
folders. Before October 2026 these files had no prefix.

### A note on `Posting_Date`

The old report put either a date or the text `"Not delivered"` in one column. A
mixed-type column cannot be stored in Parquet and confuses Power BI's type
inference, so it is two columns here: `Posting_Date` (a real nullable date) and
`Delivery_Status` (`Delivered` / `Not delivered`).

---

## Configuration

Everything environment-specific is in `.env`; every business rule is a named
constant in `config.py`.

| Setting | Default | Effect |
|---|---|---|
| `BC_USERNAME` / `BC_ACCESS_KEY` | — | credentials (Web Service Access Key) |
| `BC_BASE_URL` / `BC_COMPANY` | see `.env.example` | OData root and company |
| `BC_TIMEOUT` | 180 | seconds per HTTP request |
| `OUTPUT_DIR` | `./output` | where the Excel workbook goes |
| `PARQUET_DIR` | `OUTPUT_DIR/parquet` | where Power BI reads from |
| `INCLUDE_OPEN_ORDERS` | **true** | include live, not-yet-archived lines |
| `KEEP_HISTORY` | true | keep a dated Parquet snapshot per run |
| `SNOP_REQUIRE_ELAPSED_DEADLINE` | true | restrict `snop` to elapsed deadlines |
| `LOOKBACK_DAYS` | blank | only pull orders requested within N days |
| `SKIP_ARCHIVE_LINES` | false | escape hatch when the archive endpoint is down |

### CLI flags

```bash
.venv\Scripts\python main.py [flags]
```

| Flag | Effect |
|---|---|
| `--no-excel` | Parquet only |
| `--no-parquet` | Excel only |
| `--include-open-orders` | force open orders on for this run |
| `--no-archive-lines` | skip the archive entity (implies open orders) |
| `--lookback DAYS` | only pull orders requested within N days |
| `--no-history` | skip the dated Parquet snapshot |
| `--output DIR` | write everything somewhere else (safe for dry runs) |
| `-v` | debug logging, including per-page row counts |

`--output` redirects both Excel and Parquet, so it is the safe way to try a
change without touching production.

Logs land in `logs/otif_<date>.log` with the headline percentages, so a scheduled
run is auditable without opening any output.

---

## The two reports

### 1. `main.py` — the OTIF report

The deadline clock starts at the **PO date**. This is the source of truth for
OTIF percentages.

### 2. `release_otif.py` — the release-date report

Identical in every respect except two:

- The deadline is anchored on the order's **release date** — when it was actually
  released to fulfilment, which is when the factory's clock genuinely starts —
  rather than the PO date.
- Lead days come from `reference/lead_days.xlsx`, which the business edits in
  Excel: defaults per MTA flag, plus optional per-item and per-brand overrides.
  Item beats brand, brand beats default. `reference/lead_days.csv` is a fallback
  if the workbook is missing or locked.

```bash
.venv\Scripts\python release_otif.py                    # full run
.venv\Scripts\python release_otif.py --from 2026-04-01  # trim the window
.venv\Scripts\python release_otif.py --init-lead-days   # rebuild the lead-days workbook
```

It writes `release_*.parquet` beside the main report's files and
`OTIF Report (Release Date).xlsx`, and never modifies the main report's output.

**Two limits to understand before quoting its numbers.** `OrderReleasedDate` is
only populated in BC from **March 2026** onward, so nothing older can appear. And
the release date is read from `PostedSalesInvoiceBi`, which means **an order only
enters this report once it has been invoiced**. Late, disputed and undelivered
orders are exactly the ones not yet invoiced, so they are structurally absent and
this report reads considerably higher than the main one. Treat it as lead-time
analysis on completed business, not as the OTIF figure.

---

## Constraints and gotchas

### A closed month must not move

Once a month's deadlines have passed its OTIF result is decided — a delivery made
in August cannot rescue a July line. For that to hold in practice, the July
population has to be **complete on 31 July**, and that is what
`INCLUDE_OPEN_ORDERS=true` buys you. **Leave it on.**

With it off, the spine is the archive only, and BC archives a sales order only
when it **closes**. An order that blew its July deadline but was still open would
not exist in the extract at all — not in July's bucket, not in August's, nowhere
— and would drop into July weeks later, as a failure, when someone finally closed
it. Since the orders that miss their deadline are by definition the ones that
stay open longest, every month published too high and decayed toward the truth
over the following weeks. It is a systematic optimistic bias, not noise.

Two supporting details make it work:

- `PO_Date` falls back to the **live** order header when a document has no
  archive header yet. Without that fallback an open line has no PO date, hence no
  `Final_Deadline`, hence no deadline month at all — it drops out of every bucket
  silently, which is the leak itself. The archive header still wins wherever it
  exists, so no archived line is ever restated.
- Open and archived copies of a line are collapsed as described under
  [Row identity](#row-identity).

**The remaining ways a closed month can still change**, all rare but real:

1. An open order is **deleted** in BC rather than closed — its lines vanish.
   Arguably correct: a cancelled order is not a missed delivery.
2. **Editing `reference/mta_list.csv`.** The MTA flag drives lead days, which
   drive every deadline, applied to *all* history — not just new lines. Flipping
   one item's flag silently restates every past month it appears in. The same
   goes for `lead_days.xlsx` on the release report. This is the one most likely
   to catch you out.
3. A new order entered with a **backdated** requested delivery date, landing in a
   month already published.

### Blank dates

Business Central returns `0001-01-01` for an empty date, not null.
`extract.coerce_dates` maps that sentinel to null so the intended "fall back to
PO Date" branches actually fire.

Two situations look alike but are not, and `transform.build_otif_lines` tracks
them separately:

- **No order header at all** — no usable deadline, so the line can never be OTIF.
- **A header exists but its PO date is the sentinel** — the requested shipment
  date is still available, so a deadline can be derived.

Treating them alike wrongly scores delivered lines as failures. `selftest.py`
covers both paths.

### Pagination is not optional

The service pages at 100,000 rows and several entities exceed that unfiltered.
`bc_client` follows `@odata.nextLink` to exhaustion. It also never sends `$top`
— in OData `$top` caps the *whole* result set, not the page size, so sending it
would silently truncate.

Row filters are pushed server-side (`FILTER_*` in `config.py`), which keeps every
entity inside a single page and cuts the transfer by roughly 60%.

### The archive endpoint is flaky

`ArchiveSalesOrdeLinesBi` has been observed timing out repeatedly at 120s and
180s — even for a single row — then serving the same request in 0.1s minutes
later. It is transient server-side load, not a broken entity. See
[Troubleshooting](#troubleshooting).

### Scrap is excluded

Documents beginning `SCRAP`/`SCR`, items beginning `SCRAP`, and brands beginning
`SCRAP` are dropped (`SCRAP_*` in `config.py`).

---

## Scheduling

There are **two** schedules and the order matters:

```
09:45  Task Scheduler   ->  python main.py  ->  writes Parquet to OneDrive  (~100s)
       OneDrive syncs to SharePoint (seconds, but only while logged on)
10:45  Power BI Service ->  scheduled refresh reads the Parquet
```

Leave a comfortable gap. If Power BI refreshes before OneDrive has finished
syncing it silently reads yesterday's file, which still looks like a successful
refresh. That is what `last_run.json` and the `Refreshed_At` card are for.

Register the job:

```powershell
powershell -ExecutionPolicy Bypass -File .\register_task.ps1
```

Verify it:

```powershell
Start-ScheduledTask -TaskName "OTIF Daily Build"
Get-ScheduledTaskInfo -TaskName "OTIF Daily Build"   # LastTaskResult 0 = OK
```

`LastTaskResult` mirrors the exit code: 0 success, 1 config, 2 extraction, 3
unexpected. `run_otif.cmd` appends any non-zero exit to `logs\failures.log`.

### The one scenario that produces stale data

The trigger only fires if the account is **already logged on**. `StartWhenAvailable`
catches a missed run, but Windows launches missed tasks up to ~10 minutes after
the session becomes available and the run itself takes ~100 seconds. Signing in
at 10:10 can mean fresh Parquet at ~10:13, still syncing — and a 10:15 Power BI
refresh loses the race.

Either refresh Power BI late enough to absorb that (10:45 rather than 10:15), or
stay logged on overnight.

"Run whether user is logged on or not" does **not** fix this while `PARQUET_DIR`
points at OneDrive: the OneDrive sync client only runs inside a logged-on
session, so a logged-off run writes files locally that never reach SharePoint.
Truly login-independent scheduling means moving the destination off OneDrive — a
SharePoint library written via Graph API, a file share a gateway can read, or an
always-on host.

---

## Power BI

**Recommended: Parquet on SharePoint, refreshed with OAuth, no gateway.**

Point `PARQUET_DIR` at a folder inside a SharePoint document library synced to
this PC. Power BI Service then reads the Parquet directly using your
organisational account. Business Central credentials stay in `.env` on this
machine and never enter Power BI, and no gateway is involved.

Prefer a real SharePoint/Teams library over personal OneDrive — personal OneDrive
works, but the report then depends on one person's account surviving.

```m
let
    Source   = SharePoint.Files(
                   "https://nutasteworld-my.sharepoint.com/personal/sudhanshu_tiwari_nutasteworld_com",
                   [ApiVersion = 15]),
    Filtered = Table.SelectRows(Source, each [Name] = "PO_snop.parquet"
                                          and Text.Contains([Folder Path], "/OTIF/current/")),
    Data     = Parquet.Document(Filtered{0}[Content])
in
    Data
```

Duplicate that query per table, changing the file name. On a shared library the
URL becomes `https://nutasteworld.sharepoint.com/sites/<site-name>` and the
folder filter changes accordingly.

Point the OTIF percentage visuals at `PO_snop` and the expediting visuals at
`PO_pending` (or the `release_` equivalents). For trending, point a second query
at `history/PO_otif_lines/` and
combine the folder — each run keeps a dated snapshot with a `Snapshot_Date`
column.

Then in the Service: Dataset → Settings → Data source credentials →
**OAuth2 / organizational account**, and schedule the refresh after the Python
job.

**Do not point Power BI at Business Central directly.** That forces basic auth
plus a gateway and re-creates the original problem: unfiltered full-table pulls
on Microsoft's refresh schedule.

---

## Knowing when it last refreshed

**The log** — `logs\otif_<date>.log`. Every line carries a full date and time,
and the last line of a run is explicit:

```
2026-08-26 11:42:21  INFO    Refreshed at 2026-08-26 11:42:21 (success, 93.7s)
```

**`last_run.json`** — written beside the Parquet on **every** run, success or
failure, so a stale report is distinguishable from an unchanged one. It carries
`status`, `exit_code`, timings and the headline counts. `run_history.jsonl`
appends one line per run, so a job that quietly stopped firing shows up as a gap
in the dates.

**A `Refreshed_At` card in Power BI** — the most useful of the three, because a
chart of OTIF percentages will never reveal that the pipeline stopped:

```
Data as of = "Data as of " & FORMAT(MAX('snop'[Refreshed_At]), "dd MMM yyyy HH:mm")
```

---

## Troubleshooting

### Check the connection first

```bash
.venv\Scripts\python probe.py
```

One `$top=1` request per entity, with timings. Separates "credentials wrong" (a
fast 401) from "server slow" (a read timeout) in seconds.

### When the archive endpoint stalls

In order of preference:

1. Let the retries work — `bc_client` retries four times with exponential backoff.
2. Raise `BC_TIMEOUT` in `.env` (300 is reasonable).
3. Narrow the pull with `--lookback 540` so the server does less work.
4. Last resort: `--no-archive-lines` builds from live order lines alone. **This
   changes the population drastically** — roughly 2,300 lines instead of 42,000 —
   so treat the output as a smoke test, not a report.

`probe_archive.py` characterises that endpoint specifically, trying six request
shapes from `$top=1` up to a full pull.

### The numbers look wrong

1. Run `selftest.py` — it verifies the business rules offline against fixtures.
2. Check the run log for the warning counts: lines with no order header, blank PO
   dates, unmapped pincodes, duplicate reference rows.
3. Compare against a dated snapshot in `history/` to see exactly what changed.
4. `validate_vs_workbook.py` reconciles line-by-line against the old workbook's
   cached values, if you still need to tie back to it:

   ```bash
   .venv\Scripts\python validate_vs_workbook.py "C:\path\to\OTIF - New Copy.xlsx"
   ```

Re-run the selftest whenever a business rule in `config.py` changes.

---

## Repository layout

```
main.py                   the OTIF report: orchestration, CLI, logging, summary
release_otif.py           the release-date report (standalone, second output)

config.py                 settings, business rules, $select and $filter clauses
bc_client.py              basic auth, retries, @odata.nextLink pagination
extract.py                BC entities + reference CSVs -> typed frames
transform.py              the OTIF calculation
outputs.py                Parquet and values-only Excel writers

selftest.py               offline verification of the business rules
probe.py                  connection / latency check, one row per entity
probe_archive.py          deep probe of the slow archive endpoint
validate_vs_workbook.py   line-level reconciliation against the old .xlsx

register_task.ps1         registers the daily scheduled task
run_otif.cmd              what the scheduled task actually invokes

reference/                mta_list.csv, pincode_mapping.csv, lead_days.xlsx|csv
logs/                     one log per day, plus failures.log
output/                   Excel workbooks (git-ignored)
```

`reference/` is hand-maintained and has no source in Business Central. Changing
`mta_list.csv` or `lead_days.xlsx` restates historical months — see
[Constraints and gotchas](#constraints-and-gotchas).
