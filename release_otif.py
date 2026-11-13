"""Release-date OTIF report -- a second, standalone report.

This is deliberately separate from `main.py`. It reads the same source data and
imports the same extract/transform helpers, but it changes two rules and writes
to its own files. Nothing in the existing report is modified or overwritten.

What is different
-----------------
1. **Deadline anchor.** The existing report anchors the Internal Deadline on the
   archive header's PO Date. This one anchors it on the order's *release date* --
   the moment the order was actually released to fulfilment, which is when the
   factory's clock genuinely starts.

   `OrderReleasedDate` is not on the archive header, so it is taken from
   `PostedSalesInvoiceBi` and joined back on `Order_No`.

2. **Lead days.** The existing report hardcodes 7 days for MTA items and 14 for
   everything else. Here the numbers come from `reference/lead_days.xlsx`, which
   the business edits in Excel: defaults per MTA flag, plus optional per-item and
   per-brand overrides. `reference/lead_days.csv` is read as a fallback if the
   workbook is missing or cannot be opened.

Reporting window
----------------
`OrderReleasedDate` was switched on in BC in **March 2026** -- measured
2026-08-18, it is 0% populated for every month from 2025-05 to 2026-02, 57% in
2026-03, and ~100% from 2026-04 onward. This report therefore covers only orders
that actually carry a release date; everything older is excluded rather than
guessed at. Use `--from` to trim the window further.

Usage
-----
    python release_otif.py                    # full run: Parquet + Excel
    python release_otif.py --from 2026-04-01  # only the fully-populated era
    python release_otif.py --no-excel         # Parquet only
    python release_otif.py --init-lead-days   # (re)create the lead-days workbook
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd

import config
import extract
import outputs
import transform
from bc_client import BCClient, BusinessCentralError

log = logging.getLogger("otif.release")

EXIT_OK = 0
EXIT_CONFIG = 1
EXIT_EXTRACT = 2
EXIT_UNEXPECTED = 3

# --------------------------------------------------------------------------
# Settings local to this report (config.py is intentionally left untouched)
# --------------------------------------------------------------------------
ENTITY_INVOICE_HEADERS = "PostedSalesInvoiceBi"
SELECT_INVOICE_HEADERS = ["No", "Order_No", "OrderReleasedDate", "Posting_Date"]

# Lead days can be maintained in either file. The workbook is preferred because
# it is what the business actually edits; the CSV stays as a fallback so the
# report still runs if the workbook is missing, locked by Excel, or corrupted.
LEAD_DAYS_XLSX = config.REFERENCE_DIR / "lead_days.xlsx"
LEAD_DAYS_CSV = config.REFERENCE_DIR / "lead_days.csv"
EXCEL_NAME = "OTIF Report (Release Date).xlsx"

SCOPES = ("Default", "Item", "Brand")

# Frame names are prefixed so the Parquet files and their dated history sit
# beside the existing report's without colliding with it.
PARQUET_PREFIX = "release_"

# Fallbacks used only if lead_days.csv is missing its Default rows.
FALLBACK_LEAD_DAYS = {"MTA": 7, "NON_MTA": 14}


# --------------------------------------------------------------------------
# Extract: order release dates, via the posted sales invoice header
# --------------------------------------------------------------------------
def release_dates(client: BCClient) -> pd.DataFrame:
    """Return one release date per sales order, from PostedSalesInvoiceBi.

    An order can have several invoices, and they may carry different release
    dates. The earliest is taken: it is the moment the order first began
    releasing, and it yields the stricter deadline.
    """
    # No $filter here on purpose. This entity has no server-side index on
    # Order_No: `$filter=Order_No ne ''` makes it read-timeout at 180s and
    # exhaust every retry, while the unfiltered projection returns all ~14.8k
    # rows in well under a second. The equivalent filter is applied below.
    records = client.fetch(
        ENTITY_INVOICE_HEADERS,
        select=SELECT_INVOICE_HEADERS,
    )
    frame = pd.DataFrame.from_records(records)
    for col in SELECT_INVOICE_HEADERS:
        if col not in frame.columns:
            frame[col] = pd.NA
    frame = frame[SELECT_INVOICE_HEADERS].copy()

    frame["Order_No"] = frame["Order_No"].astype("string").fillna("").str.strip()
    frame = extract.coerce_dates(frame, ["OrderReleasedDate", "Posting_Date"])

    total = len(frame)
    frame = frame[(frame["Order_No"] != "") & frame["OrderReleasedDate"].notna()]
    log.info("  invoice headers: %d -> %d with an order link and a release date",
             total, len(frame))

    per_order = (
        frame.sort_values("OrderReleasedDate")
        .drop_duplicates(subset=["Order_No"], keep="first")
        .rename(columns={"OrderReleasedDate": "Release_Date"})
        [["Order_No", "Release_Date"]]
        .reset_index(drop=True)
    )
    if len(per_order):
        log.info("  release dates: %d order(s), %s .. %s",
                 len(per_order),
                 per_order["Release_Date"].min().date(),
                 per_order["Release_Date"].max().date())
    else:
        log.warning("  no order release dates found at all")
    return per_order


# --------------------------------------------------------------------------
# Extract: the editable lead-days table
# --------------------------------------------------------------------------
def _read_lead_days_source() -> tuple[pd.DataFrame | None, Path | None]:
    """Return the lead-days rows from the workbook, else the CSV, else nothing."""
    if LEAD_DAYS_XLSX.exists():
        try:
            frame = pd.read_excel(
                LEAD_DAYS_XLSX, sheet_name="Lead Days", dtype=str, engine="openpyxl"
            )
            return frame.fillna(""), LEAD_DAYS_XLSX
        except ImportError:
            log.warning("  %s exists but openpyxl is not installed "
                        "(pip install openpyxl) -- falling back to the CSV",
                        LEAD_DAYS_XLSX.name)
        except Exception as exc:  # noqa: BLE001 - reference data must not stop the run
            log.warning("  could not read %s (%s: %s) -- falling back to the CSV",
                        LEAD_DAYS_XLSX.name, type(exc).__name__, exc)

    if LEAD_DAYS_CSV.exists():
        return pd.read_csv(LEAD_DAYS_CSV, dtype=str, comment="#").fillna(""), LEAD_DAYS_CSV
    return None, None


def load_lead_days() -> dict[str, dict[str, int]]:
    """Read the lead-days table into {scope: {key: days}}.

    Prefers reference/lead_days.xlsx, falls back to reference/lead_days.csv.
    Missing file or missing rows fall back to the built-in 7/14, so the report
    still runs rather than failing on a reference-data problem.
    """
    table: dict[str, dict[str, int]] = {"Default": {}, "Item": {}, "Brand": {}}

    frame, source = _read_lead_days_source()
    if frame is None:
        log.warning("  no lead-days file found (looked for %s and %s) -- "
                    "using built-in defaults MTA=%d, non-MTA=%d",
                    LEAD_DAYS_XLSX.name, LEAD_DAYS_CSV.name,
                    FALLBACK_LEAD_DAYS["MTA"], FALLBACK_LEAD_DAYS["NON_MTA"])
        table["Default"] = dict(FALLBACK_LEAD_DAYS)
        return table

    log.info("  lead days source: %s", source.name)
    for col in ("Scope", "Key", "Lead_Days"):
        if col not in frame.columns:
            raise ValueError(f"{source} is missing the {col!r} column")
        frame[col] = frame[col].astype("string").fillna("").str.strip()

    days = pd.to_numeric(frame["Lead_Days"], errors="coerce")
    bad = frame[days.isna() & (frame["Lead_Days"] != "")]
    for _, row in bad.iterrows():
        log.warning("  %s: ignoring %s/%s -- %r is not a number",
                    source.name, row["Scope"], row["Key"], row["Lead_Days"])

    for (scope, key), value in zip(zip(frame["Scope"], frame["Key"]), days):
        scope = scope.strip().title()
        if pd.isna(value) or scope not in table or not key:
            continue
        table[scope][key.upper() if scope == "Default" else key] = int(value)

    for key, fallback in FALLBACK_LEAD_DAYS.items():
        if key not in table["Default"]:
            log.warning("  %s has no Default/%s row -- using %d",
                        source.name, key, fallback)
            table["Default"][key] = fallback

    log.info("  lead days: MTA=%d, non-MTA=%d, %d item override(s), %d brand override(s)",
             table["Default"]["MTA"], table["Default"]["NON_MTA"],
             len(table["Item"]), len(table["Brand"]))
    return table


def write_lead_days_workbook(path: Path | None = None) -> Path:
    """Create reference/lead_days.xlsx, seeded from whatever is configured now.

    Built with dropdowns on Scope and a whole-number rule on Lead_Days, so the
    business can edit it without producing values the loader has to reject. Run
    via `python release_otif.py --init-lead-days`.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    path = path or LEAD_DAYS_XLSX
    table = load_lead_days()

    book = Workbook()
    sheet = book.active
    sheet.title = "Lead Days"

    header_fill = PatternFill("solid", start_color="1F3864")
    header_font = Font(bold=True, color="FFFFFF")
    for col, name in enumerate(("Scope", "Key", "Lead_Days"), start=1):
        cell = sheet.cell(row=1, column=col, value=name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left")

    rows = [("Default", key, table["Default"][key]) for key in ("MTA", "NON_MTA")]
    rows += [("Item", k, v) for k, v in sorted(table["Item"].items())]
    rows += [("Brand", k, v) for k, v in sorted(table["Brand"].items())]
    for r, (scope, key, days) in enumerate(rows, start=2):
        sheet.cell(row=r, column=1, value=scope)
        sheet.cell(row=r, column=2, value=key)
        sheet.cell(row=r, column=3, value=int(days))

    # Room to add overrides by typing, with the rules already in place.
    last_row = 400
    scope_rule = DataValidation(
        type="list", formula1=f'"{",".join(SCOPES)}"', allow_blank=True,
        showDropDown=False, errorTitle="Pick a scope",
        error=f"Scope must be one of: {', '.join(SCOPES)}.",
    )
    days_rule = DataValidation(
        type="whole", operator="between", formula1=0, formula2=365,
        allow_blank=True, errorTitle="Lead days",
        error="Lead days must be a whole number of days between 0 and 365.",
    )
    sheet.add_data_validation(scope_rule)
    sheet.add_data_validation(days_rule)
    scope_rule.add(f"A2:A{last_row}")
    days_rule.add(f"C2:C{last_row}")

    for col, width in ((1, 14), (2, 30), (3, 12)):
        sheet.column_dimensions[get_column_letter(col)].width = width
    sheet.freeze_panes = "A2"

    guide = book.create_sheet("How to use")
    guide.column_dimensions["A"].width = 104
    lines = [
        ("Lead days for the release-date OTIF report", True),
        ("", False),
        ("Internal Deadline = Order Release Date + Lead_Days", False),
        ("", False),
        ("The Lead Days sheet controls how many days each item gets after its order", False),
        ("is released to fulfilment. Edit the numbers, save, and re-run the report:", False),
        ("    python release_otif.py", False),
        ("", False),
        ("Scope column", True),
        ("  Default   Key is MTA or NON_MTA. The number used when an item has no override.", False),
        ("            Whether an item is MTA comes from reference/mta_list.csv.", False),
        ("  Item      Key is an item number, e.g. 112-570. Applies to that item only.", False),
        ("  Brand     Key is a brand name, e.g. CHANNEL. Applies to that brand.", False),
        ("", False),
        ("Precedence", True),
        ("  Item beats Brand, and Brand beats Default. Both Default rows must exist;", False),
        ("  if one is missing the report warns and falls back to 7 (MTA) / 14 (non-MTA).", False),
        ("", False),
        ("Worked example", True),
        ("  An order released 31-Mar-2026 containing an MTA item at 7 lead days gets an", False),
        ("  internal deadline of 07-Apr-2026. The same order with a non-MTA item at 14", False),
        ("  days gets 14-Apr-2026.", False),
        ("", False),
        ("Notes", True),
        ("  Rows with a blank or non-numeric Lead_Days are skipped, and the report logs", False),
        ("  which ones. Only this workbook is read; reference/lead_days.csv is a fallback", False),
        ("  used if this file is missing or cannot be opened. Do not rename the sheet.", False),
    ]
    for r, (text, bold) in enumerate(lines, start=1):
        cell = guide.cell(row=r, column=1, value=text)
        if bold:
            cell.font = Font(bold=True)

    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    log.info("wrote %s (%d lead-day row(s), editable in Excel)", path, len(rows))
    return path


def resolve_lead_days(frame: pd.DataFrame, table: dict[str, dict[str, int]]) -> pd.Series:
    """Lead days per line: Item override -> Brand override -> MTA default."""
    default = np.where(
        frame["MTA"], table["Default"]["MTA"], table["Default"]["NON_MTA"]
    )
    days = pd.Series(default, index=frame.index, dtype="float")

    if table["Brand"] and "Brand" in frame.columns:
        by_brand = frame["Brand"].map(table["Brand"])
        days = by_brand.astype("float").fillna(days)

    if table["Item"]:
        by_item = frame["No"].map(table["Item"])
        days = by_item.astype("float").fillna(days)

    return days.round().astype(int)


# --------------------------------------------------------------------------
# Transform: re-anchor the deadline on the release date
# --------------------------------------------------------------------------
def build_release_otif(
    otif_lines: pd.DataFrame,
    ship_lines: pd.DataFrame,
    releases: pd.DataFrame,
    lead_table: dict[str, dict[str, int]],
    release_from: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Take the standard OTIF lines and recompute the deadline off release date.

    The base frame is produced by the existing `transform.build_otif_lines`, so
    every other column -- keys, quantities, brand, city, delivery status -- keeps
    exactly the same meaning as in the main report. Only the deadline and the
    flags derived from it are recomputed here.
    """
    frame = otif_lines.copy()

    # --- attach the release date -----------------------------------------
    release_map = releases.set_index("Order_No")["Release_Date"]
    frame["Release_Date"] = frame["Document_No"].map(release_map)

    before = len(frame)
    frame = frame[frame["Release_Date"].notna()].copy()
    log.info("  release window: %d of %d line(s) have an order release date "
             "(%d excluded as pre-release-date)",
             len(frame), before, before - len(frame))
    if release_from is not None:
        kept = frame["Release_Date"] >= release_from
        log.info("  --from %s: dropping %d line(s) released earlier",
                 release_from.date(), int((~kept).sum()))
        frame = frame[kept].copy()
    if frame.empty:
        log.warning("  no lines left after the release-date filter")
        return frame

    # --- lead days --------------------------------------------------------
    frame["Lead_Days"] = resolve_lead_days(frame, lead_table)

    # --- internal deadline = release date + lead days ---------------------
    frame["Internal_Deadline"] = frame["Release_Date"] + pd.to_timedelta(
        frame["Lead_Days"], unit="D"
    )

    # --- requested shipment date ------------------------------------------
    # Same rule as the main report, except the blank-requested-date fallback is
    # the release date rather than the PO Date, so both halves of the MAX below
    # are anchored consistently.
    shifted = frame["Requested_Delivery_Date"] - pd.Timedelta(
        days=config.SHIPMENT_LEAD_OFFSET_DAYS
    )
    frame["Requested_Shipment_Date"] = shifted.where(
        frame["Requested_Delivery_Date"].notna(), frame["Release_Date"]
    )

    # --- final deadline ----------------------------------------------------
    # MAX of the two, as in the main report. Every row here has a release date,
    # so the Internal Deadline is always present and Final_Deadline is never null.
    frame["Final_Deadline"] = frame[
        ["Internal_Deadline", "Requested_Shipment_Date"]
    ].max(axis=1)

    # --- redo the quantity sum against the new deadline -------------------
    # Reuses the main report's SUMIFS replacement so the two reports cannot
    # drift apart on what "delivered before the deadline" means.
    frame["Qty_Delivered_Before_Deadline"] = transform._qty_before_deadline(
        frame, ship_lines
    )

    frame["In_Full"] = (
        frame["Qty_Delivered_Before_Deadline"] >= frame["Quantity"].fillna(0)
    )
    frame["On_Time"] = frame["In_Full"]
    frame["OTIF"] = frame["In_Full"]
    return frame


# --------------------------------------------------------------------------
# Output shape
# --------------------------------------------------------------------------
RELEASE_OTIF_COLUMNS = [
    "Document_No", "Line_No", "No", "Description", "Quantity", "Line_Amount",
    "Qty_to_Ship", "Requested_Delivery_Date", "Key", "Key_Req",
    "PO_Date", "Release_Date", "MTA", "Lead_Days",
    "Internal_Deadline", "Requested_Shipment_Date", "Final_Deadline",
    "Deadline_Status", "Days_To_Deadline",
    "Qty_Delivered_Before_Deadline", "Qty_Short", "Posting_Date",
    "Delivery_Status", "In_Full", "On_Time", "OTIF",
    "Business_Vertical", "Brand", "City", "Customer_Name", "Source",
]

RELEASE_SNOP_COLUMNS = [
    "Key_Req", "Document_No", "Customer_Name", "Brand", "Business_Vertical",
    "No", "Description", "PO_Date", "Release_Date", "Requested_Delivery_Date",
    "MTA", "Lead_Days", "Internal_Deadline", "Requested_Shipment_Date",
    "Final_Deadline", "Deadline_Status", "Qty_Delivered_Before_Deadline",
    "Qty_Short", "Posting_Date", "Delivery_Status", "In_Full", "On_Time",
    "OTIF", "Line_Amount", "City", "Source",
]

RELEASE_PENDING_COLUMNS = [
    "Final_Deadline", "Days_To_Deadline", "Document_No", "Line_No", "No",
    "Description", "Customer_Name", "Brand", "Business_Vertical", "City",
    "Quantity", "Qty_Delivered_Before_Deadline", "Qty_Short",
    "Delivery_Status", "In_Full", "MTA", "Lead_Days", "Release_Date",
    "Requested_Delivery_Date", "Line_Amount", "Key_Req", "Source",
]


# --------------------------------------------------------------------------
# Runner
# --------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build the release-date OTIF report from Business Central."
    )
    parser.add_argument("--from", dest="release_from", metavar="YYYY-MM-DD",
                        help="only orders released on or after this date "
                             "(default: every order that has a release date)")
    parser.add_argument("--init-lead-days", action="store_true",
                        help="(re)create reference/lead_days.xlsx from the current "
                             "settings and exit, without running the report")
    parser.add_argument("--no-excel", action="store_true", help="skip the Excel workbook")
    parser.add_argument("--no-parquet", action="store_true", help="skip the Parquet files")
    parser.add_argument("--no-history", action="store_true",
                        help="do not keep a dated Parquet snapshot")
    parser.add_argument("--include-open-orders", action="store_true", default=None,
                        help="include live (not yet archived) sales order lines")
    parser.add_argument("--lookback", type=int, metavar="DAYS",
                        help="only pull order lines requested within this many days")
    parser.add_argument("--output", type=Path, metavar="DIR",
                        help="override the output directory")
    parser.add_argument("-v", "--verbose", action="store_true", help="debug logging")
    return parser.parse_args(argv)


def run(args: argparse.Namespace, run_date: date, stats: dict) -> int:
    # Reference-data maintenance only -- no BC connection needed.
    if args.init_lead_days:
        write_lead_days_workbook()
        return EXIT_OK

    problems = config.validate()
    if problems:
        for problem in problems:
            log.error("config: %s", problem)
        return EXIT_CONFIG

    if args.lookback is not None:
        config.LOOKBACK_DAYS = args.lookback
    include_open = (
        config.INCLUDE_OPEN_ORDERS
        if args.include_open_orders is None
        else args.include_open_orders
    )
    release_from = (
        pd.Timestamp(args.release_from) if args.release_from else None
    )

    output_dir = args.output or config.OUTPUT_DIR
    parquet_dir = config.PARQUET_DIR if args.output is None else output_dir / "parquet"
    refreshed_at = datetime.now()
    stats.update({
        "report": "release_otif",
        "parquet_dir": str(parquet_dir),
        "started_at": refreshed_at,
        "release_from": args.release_from,
    })

    log.info("Release-date OTIF run %s", run_date.isoformat())
    log.info("  started at          : %s", refreshed_at.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("  endpoint            : %s", config.BC_BASE_URL)
    log.info("  company             : %s", config.BC_COMPANY)
    log.info("  release date from   : %s", args.release_from or "any")
    log.info("  include open orders : %s", include_open)

    # ---------------- extract ----------------
    log.info("Extracting from Business Central...")
    started = time.perf_counter()
    try:
        with BCClient() as client:
            line_frames: list[pd.DataFrame] = []
            if not config.SKIP_ARCHIVE_LINES:
                line_frames.append(extract.archive_order_lines(client))
            if include_open:
                line_frames.append(extract.open_order_lines(client))
            if not line_frames:
                log.error("no order line source enabled")
                return EXIT_CONFIG
            order_lines = (
                line_frames[0] if len(line_frames) == 1
                else pd.concat(line_frames, ignore_index=True)
            )

            ship_lines_raw = extract.shipment_lines(client)
            ship_headers = extract.shipment_headers(client)
            combined_headers, archive_headers = extract.order_headers(client)
            releases = release_dates(client)
    except BusinessCentralError as exc:
        log.error("extraction failed: %s", exc)
        return EXIT_EXTRACT

    mta = extract.mta_list()
    pincodes = extract.pincode_mapping()
    lead_table = load_lead_days()
    log.info("Extraction finished in %.1fs", time.perf_counter() - started)

    # ---------------- transform ----------------
    log.info("Transforming...")
    started = time.perf_counter()
    combined_headers = transform.add_city(combined_headers, pincodes)
    ship_lines = transform.prepare_shipment_lines(ship_lines_raw, ship_headers)

    # Base frame from the existing pipeline, then re-anchored on release date.
    base = transform.build_otif_lines(
        order_lines=order_lines,
        ship_lines=ship_lines,
        archive_headers=archive_headers,
        combined_headers=combined_headers,
        mta=mta,
    )
    otif_lines = build_release_otif(
        base, ship_lines, releases, lead_table, release_from=release_from
    )
    if otif_lines.empty:
        log.error("no lines carry an order release date -- nothing to report")
        return EXIT_EXTRACT

    otif_lines = transform.add_deadline_status(otif_lines, pd.Timestamp(run_date))
    cutoff = pd.Timestamp(run_date) if config.SNOP_REQUIRE_ELAPSED_DEADLINE else None
    snop = transform.build_snop(otif_lines, deadline_cutoff=cutoff)
    pending = transform.build_pending(otif_lines)
    log.info("Transform finished in %.1fs", time.perf_counter() - started)

    otif_out = transform.select_columns(otif_lines, RELEASE_OTIF_COLUMNS)
    snop_out = transform.select_columns(snop, RELEASE_SNOP_COLUMNS)
    pending_out = transform.select_columns(pending, RELEASE_PENDING_COLUMNS)
    log_summary(otif_out, cutoff)

    assessable = otif_out if cutoff is None else otif_out[
        ~(otif_out["Final_Deadline"] > cutoff)
    ]
    stats.update({
        "order_lines": len(otif_out),
        "snop_rows": len(snop_out),
        "pending_rows": len(pending_out),
        "assessable_lines": len(assessable),
        "otif_achieved": int(assessable["OTIF"].sum()) if len(assessable) else 0,
        "otif_pct": round(
            100.0 * int(assessable["OTIF"].sum()) / len(assessable), 2
        ) if len(assessable) else None,
        "release_date_min": str(otif_out["Release_Date"].min().date()),
        "release_date_max": str(otif_out["Release_Date"].max().date()),
    })

    # ---------------- load ----------------
    if not args.no_parquet:
        log.info("Writing Parquet...")
        outputs.write_parquet(
            {
                f"{PARQUET_PREFIX}otif_lines": otif_out,
                f"{PARQUET_PREFIX}snop": snop_out,
                f"{PARQUET_PREFIX}pending": pending_out,
            },
            run_date=run_date,
            refreshed_at=refreshed_at,
            parquet_dir=parquet_dir,
            keep_history=False if args.no_history else None,
        )

    if not args.no_excel:
        log.info("Writing Excel...")
        outputs.write_excel(
            {"OTIF Lines": otif_out, "SNOP": snop_out, "Pending": pending_out},
            path=output_dir / EXCEL_NAME,
        )

    log.info("Done.")
    return EXIT_OK


def log_summary(otif: pd.DataFrame, cutoff: pd.Timestamp | None) -> None:
    """Headline numbers, so a scheduled run is auditable from the log alone."""
    total = len(otif)
    if not total:
        log.warning("summary: no OTIF rows produced")
        return

    delivered = int((otif["Delivery_Status"] == "Delivered").sum())
    log.info("Summary (release-date basis)")
    log.info("  order lines      : %7d", total)
    log.info("  release window   : %s .. %s",
             otif["Release_Date"].min().date(), otif["Release_Date"].max().date())
    log.info("  delivered        : %7d (%5.1f%%)", delivered, 100.0 * delivered / total)
    log.info("  MTA items        : %7d", int(otif["MTA"].sum()))
    log.info("  lead days in use : %s",
             ", ".join(f"{int(d)}d x{n}" for d, n in
                       otif["Lead_Days"].value_counts().sort_index().items()))

    if cutoff is not None:
        pending = otif["Final_Deadline"] > cutoff
        assessable = otif[~pending]
        log.info("  not yet due      : %7d (excluded from OTIF, deadline after %s)",
                 int(pending.sum()), cutoff.date())
    else:
        assessable = otif

    n = len(assessable)
    if not n:
        log.warning("  no assessable lines -- every deadline is in the future")
        return
    achieved = int(assessable["OTIF"].sum())
    log.info("  assessable lines : %7d", n)
    log.info("  OTIF achieved    : %7d (%5.1f%% of assessable)",
             achieved, 100.0 * achieved / n)

    by_month = (
        assessable.assign(
            Month=lambda d: d["Release_Date"].dt.to_period("M").astype(str)
        )
        .groupby("Month")["OTIF"].agg(["size", "sum"])
    )
    log.info("  by release month (assessable only):")
    for month, row in by_month.iterrows():
        pct = 100.0 * row["sum"] / row["size"] if row["size"] else 0.0
        log.info("    %s  %6d lines  OTIF %5.1f%%", month, int(row["size"]), pct)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    run_date = date.today()

    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                config.LOG_DIR / f"otif_release_{run_date:%Y-%m-%d}.log",
                encoding="utf-8",
            ),
        ],
        force=True,
    )
    logging.getLogger("urllib3").setLevel(logging.WARNING)

    stats: dict = {"run_date": run_date}
    started = time.perf_counter()
    code = EXIT_UNEXPECTED
    try:
        code = run(args, run_date, stats)
    except KeyboardInterrupt:
        log.warning("interrupted")
        stats["error"] = "interrupted"
    except Exception as exc:  # noqa: BLE001 - recorded, then re-reported
        log.exception("unexpected failure")
        stats["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        finished = datetime.now()
        stats.update({
            "status": "success" if code == EXIT_OK else "failed",
            "exit_code": code,
            "finished_at": finished,
            "duration_seconds": round(time.perf_counter() - started, 1),
        })
        log.info("Refreshed at %s (%s, %.1fs)",
                 finished.strftime("%Y-%m-%d %H:%M:%S"),
                 stats["status"], stats["duration_seconds"])
    return code


if __name__ == "__main__":
    sys.exit(main())
