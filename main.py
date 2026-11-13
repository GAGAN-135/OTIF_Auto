"""OTIF pipeline entry point.

Typical use:
    python main.py                       # full run: Parquet + Excel
    python main.py --lookback 540        # only recent orders
    python main.py --include-open-orders # add live, not-yet-archived lines
    python main.py --no-excel            # Parquet only (Power BI refresh)
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import config
import extract
import outputs
import transform
from bc_client import BCClient, BusinessCentralError

log = logging.getLogger("otif")

EXIT_OK = 0
EXIT_CONFIG = 1
EXIT_EXTRACT = 2
EXIT_UNEXPECTED = 3


def setup_logging(verbose: bool, run_date: date) -> None:
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    handlers: list[logging.Handler] = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(
            config.LOG_DIR / f"otif_{run_date:%Y-%m-%d}.log", encoding="utf-8"
        ),
    ]
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s %(message)s",
        # Full date, not just the time: for a scheduled job a bare clock time
        # cannot tell you which day a line came from.
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )
    # requests/urllib3 chatter is noise at INFO.
    logging.getLogger("urllib3").setLevel(logging.WARNING)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build the OTIF report from Business Central OData."
    )
    parser.add_argument("--no-excel", action="store_true", help="skip the Excel workbook")
    parser.add_argument("--no-parquet", action="store_true", help="skip the Parquet files")
    parser.add_argument("--no-history", action="store_true",
                        help="do not keep a dated Parquet snapshot")
    parser.add_argument("--include-open-orders", action="store_true", default=None,
                        help="include live (not yet archived) sales order lines")
    parser.add_argument("--no-archive-lines", action="store_true", default=None,
                        help="skip ArchiveSalesOrdeLinesBi (use when that endpoint "
                             "times out); implies --include-open-orders")
    parser.add_argument("--lookback", type=int, metavar="DAYS",
                        help="only pull order lines requested within this many days")
    parser.add_argument("--output", type=Path, metavar="DIR",
                        help="override the output directory")
    parser.add_argument("-v", "--verbose", action="store_true", help="debug logging")
    return parser.parse_args(argv)


def run(args: argparse.Namespace, run_date: date, stats: dict) -> int:
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
    skip_archive = (
        config.SKIP_ARCHIVE_LINES
        if args.no_archive_lines is None
        else args.no_archive_lines
    )
    if skip_archive and not include_open:
        log.warning("archive lines are skipped, so open orders are being enabled -- "
                    "otherwise there would be no order lines at all")
        include_open = True
    output_dir = args.output or config.OUTPUT_DIR
    parquet_dir = config.PARQUET_DIR if args.output is None else output_dir / "parquet"
    refreshed_at = datetime.now()
    stats.update({
        "parquet_dir": str(parquet_dir),
        "started_at": refreshed_at,
        "include_open_orders": include_open,
        "skip_archive_lines": skip_archive,
    })

    log.info("OTIF run %s", run_date.isoformat())
    log.info("  started at          : %s", refreshed_at.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("  endpoint            : %s", config.BC_BASE_URL)
    log.info("  company             : %s", config.BC_COMPANY)
    log.info("  include open orders : %s", include_open)
    log.info("  skip archive lines  : %s", skip_archive)
    log.info("  request timeout     : %ds", config.REQUEST_TIMEOUT)
    log.info("  lookback days       : %s", config.LOOKBACK_DAYS or "none (full pull)")

    # ---------------- extract ----------------
    log.info("Extracting from Business Central...")
    started = time.perf_counter()
    try:
        with BCClient() as client:
            line_frames: list[pd.DataFrame] = []
            if not skip_archive:
                line_frames.append(extract.archive_order_lines(client))
            if include_open:
                line_frames.append(extract.open_order_lines(client))
            order_lines = (
                line_frames[0] if len(line_frames) == 1
                else pd.concat(line_frames, ignore_index=True)
            )

            ship_lines_raw = extract.shipment_lines(client)
            ship_headers = extract.shipment_headers(client)
            combined_headers, archive_headers = extract.order_headers(client)
    except BusinessCentralError as exc:
        log.error("extraction failed: %s", exc)
        return EXIT_EXTRACT

    mta = extract.mta_list()
    pincodes = extract.pincode_mapping()
    log.info("Extraction finished in %.1fs", time.perf_counter() - started)

    # ---------------- transform ----------------
    log.info("Transforming...")
    started = time.perf_counter()
    combined_headers = transform.add_city(combined_headers, pincodes)
    ship_lines = transform.prepare_shipment_lines(ship_lines_raw, ship_headers)
    otif_lines = transform.build_otif_lines(
        order_lines=order_lines,
        ship_lines=ship_lines,
        archive_headers=archive_headers,
        combined_headers=combined_headers,
        mta=mta,
    )
    otif_lines = transform.add_deadline_status(otif_lines, pd.Timestamp(run_date))
    cutoff = (
        pd.Timestamp(run_date) if config.SNOP_REQUIRE_ELAPSED_DEADLINE else None
    )
    snop = transform.build_snop(otif_lines, deadline_cutoff=cutoff)
    pending = transform.build_pending(otif_lines)
    log.info("Transform finished in %.1fs", time.perf_counter() - started)

    otif_out = transform.select_columns(otif_lines, transform.OTIF_COLUMNS)
    snop_out = transform.select_columns(snop, transform.SNOP_COLUMNS)
    pending_out = transform.select_columns(pending, transform.PENDING_COLUMNS)
    log_summary(otif_out, cutoff)
    log_pending(pending_out)

    assessable = otif_out if cutoff is None else otif_out[
        ~(otif_out["Final_Deadline"] > cutoff)
    ]
    stats.update({
        "order_lines": len(otif_out),
        "snop_rows": len(snop_out),
        "pending_rows": len(pending_out),
        "pending_incomplete": int((~pending_out["In_Full"]).sum()) if len(pending_out) else 0,
        "assessable_lines": len(assessable),
        "otif_achieved": int(assessable["OTIF"].sum()) if len(assessable) else 0,
        "otif_pct": round(
            100.0 * int(assessable["OTIF"].sum()) / len(assessable), 2
        ) if len(assessable) else None,
    })

    # ---------------- load ----------------
    if not args.no_parquet:
        log.info("Writing Parquet...")
        prefix = config.PARQUET_PREFIX
        outputs.write_parquet(
            {
                f"{prefix}otif_lines": otif_out,
                f"{prefix}snop": snop_out,
                f"{prefix}pending": pending_out,
                f"{prefix}sales_orders": combined_headers,
                f"{prefix}shipment_lines": ship_lines,
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
            path=output_dir / config.EXCEL_NAME,
        )

    log.info("Done.")
    return EXIT_OK


def log_summary(otif: pd.DataFrame, cutoff: pd.Timestamp | None) -> None:
    """Log the headline numbers so a scheduled run is auditable from the log."""
    total = len(otif)
    if not total:
        log.warning("summary: no OTIF rows produced")
        return

    delivered = int((otif["Delivery_Status"] == "Delivered").sum())
    log.info("Summary")
    log.info("  order lines      : %7d", total)
    log.info("  delivered        : %7d (%5.1f%%)", delivered, 100.0 * delivered / total)
    log.info("  MTA items        : %7d", int(otif["MTA"].sum()))

    # OTIF is only meaningful once the deadline has passed. Reporting it over
    # every line would count not-yet-due orders as failures.
    if cutoff is not None:
        pending = otif["Final_Deadline"] > cutoff
        assessable = otif[~pending]
        log.info("  not yet due      : %7d (excluded from OTIF, deadline after %s)",
                 int(pending.sum()), cutoff.date())
    else:
        assessable = otif
        log.info("  not yet due      :       0 (no deadline cutoff applied)")

    n = len(assessable)
    if not n:
        log.warning("  no assessable lines -- every deadline is in the future")
        return
    achieved = int(assessable["OTIF"].sum())
    log.info("  assessable lines : %7d", n)
    log.info("  OTIF achieved    : %7d (%5.1f%% of assessable)", achieved,
             100.0 * achieved / n)

    if assessable["PO_Date"].notna().any():
        by_month = (
            assessable.dropna(subset=["PO_Date"])
            .assign(Month=lambda d: d["PO_Date"].dt.to_period("M").astype(str))
            .groupby("Month")["OTIF"]
            .agg(["size", "sum"])
            .tail(6)
        )
        log.info("  last 6 months by PO month (assessable only):")
        for month, row in by_month.iterrows():
            pct = 100.0 * row["sum"] / row["size"] if row["size"] else 0.0
            log.info("    %s  %6d lines  OTIF %5.1f%%", month, int(row["size"]), pct)


def log_pending(pending: pd.DataFrame) -> None:
    """Log the forward order book -- the actionable half of the report."""
    if pending.empty:
        log.info("Pending (deadline not yet reached): none")
        return

    incomplete = pending[~pending["In_Full"]]
    log.info("Pending (deadline not yet reached)")
    log.info("  lines            : %7d", len(pending))
    log.info("  not yet complete : %7d  (%.0f units short)",
             len(incomplete), incomplete["Qty_Short"].sum())
    log.info("  deadline range   : %s .. %s",
             pending["Final_Deadline"].min().date(),
             pending["Final_Deadline"].max().date())

    by_week = (
        incomplete.groupby(incomplete["Final_Deadline"].dt.to_period("W").astype(str))
        .agg(lines=("Qty_Short", "size"), short=("Qty_Short", "sum"))
        .head(4)
    )
    if len(by_week):
        log.info("  soonest shortfalls by deadline week:")
        for week, row in by_week.iterrows():
            log.info("    %s  %5d line(s)  %8.0f units short",
                     week, int(row["lines"]), row["short"])


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    run_date = date.today()
    setup_logging(args.verbose, run_date)

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
        try:
            outputs.write_run_record(
                stats, parquet_dir=Path(stats["parquet_dir"])
                if stats.get("parquet_dir") else None,
            )
        except Exception:  # noqa: BLE001 - never let bookkeeping mask the result
            log.exception("could not write the run record")
        log.info("Refreshed at %s (%s, %.1fs)",
                 finished.strftime("%Y-%m-%d %H:%M:%S"),
                 stats["status"], stats["duration_seconds"])
    return code


if __name__ == "__main__":
    sys.exit(main())
