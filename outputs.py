"""Output writers: Parquet (for Power BI) and a values-only Excel workbook.

The Excel file contains no formulas, no query connections, and no hidden raw
data sheets. That is the entire reason it opens instantly instead of crashing:
the old workbook carried 1.16M live formulas, none of them shared, each doing a
linear scan over a 15k-37k row table.
"""

from __future__ import annotations

import json
import logging
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import xlsxwriter

import config

log = logging.getLogger(__name__)

# Columns rendered with a currency-style number format in Excel.
_AMOUNT_COLUMNS = {"Line_Amount"}
_DATE_FORMAT = "dd-mmm-yyyy"


# --------------------------------------------------------------------------
# Parquet
# --------------------------------------------------------------------------
def write_run_record(
    record: dict,
    parquet_dir: Path | None = None,
) -> Path:
    """Write last_run.json next to the data: what ran, when, and whether it worked.

    Deliberately written on failure too, so a stale report is distinguishable
    from a report that simply has not changed. Power BI can load this file to
    show a "data as of ..." caption and to surface a failed overnight run.
    """
    parquet_dir = parquet_dir or config.PARQUET_DIR
    parquet_dir.mkdir(parents=True, exist_ok=True)
    target = parquet_dir / "last_run.json"
    target.write_text(json.dumps(record, indent=2, default=str), encoding="utf-8")

    # Append-only history, so a run that silently stopped firing is visible.
    history = parquet_dir / "run_history.jsonl"
    with history.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, default=str) + "\n")

    log.info("  run record: %s (status=%s)", target, record.get("status"))
    return target


def write_parquet(
    frames: dict[str, pd.DataFrame],
    run_date: date,
    refreshed_at: datetime | None = None,
    parquet_dir: Path | None = None,
    keep_history: bool | None = None,
) -> list[Path]:
    """Write each frame to <parquet_dir>/current/<name>.parquet.

    A dated copy also lands in history/ when enabled, so Power BI can trend OTIF
    over time -- something the workbook could never do, because each refresh
    overwrote the previous state.
    """
    parquet_dir = parquet_dir or config.PARQUET_DIR
    keep_history = config.KEEP_HISTORY if keep_history is None else keep_history

    current_dir = parquet_dir / "current"
    current_dir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    for name, frame in frames.items():
        stamped = frame.copy()
        # Snapshot_Date is the reporting day (use it to slice history);
        # Refreshed_At is the wall-clock moment the data was pulled (use it for
        # a "data as of ..." caption).
        stamped["Snapshot_Date"] = pd.Timestamp(run_date)
        if refreshed_at is not None:
            stamped["Refreshed_At"] = pd.Timestamp(refreshed_at)

        target = current_dir / f"{name}.parquet"
        stamped.to_parquet(target, engine="pyarrow", compression="snappy", index=False)
        written.append(target)
        log.info("  parquet: %-22s %7d rows  %6.1f KB  -> %s",
                 name, len(stamped), target.stat().st_size / 1024, target)

        if keep_history:
            hist_dir = config.HISTORY_DIR / name
            hist_dir.mkdir(parents=True, exist_ok=True)
            hist_target = hist_dir / f"{name}_{run_date:%Y-%m-%d}.parquet"
            stamped.to_parquet(hist_target, engine="pyarrow",
                               compression="snappy", index=False)
            written.append(hist_target)

    return written


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def write_excel(
    sheets: dict[str, pd.DataFrame],
    path: Path,
) -> Path:
    """Write a values-only workbook, one sheet per frame."""
    path.parent.mkdir(parents=True, exist_ok=True)

    # constant_memory keeps peak RAM flat by flushing each row as it is written,
    # which matters because these sheets are tens of thousands of rows.
    workbook = xlsxwriter.Workbook(
        str(path),
        {"constant_memory": True, "default_date_format": _DATE_FORMAT},
    )
    try:
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1F3864", "font_color": "white",
            "border": 1, "align": "left", "valign": "vcenter", "text_wrap": False,
        })
        date_fmt = workbook.add_format({"num_format": _DATE_FORMAT})
        amount_fmt = workbook.add_format({"num_format": "#,##0.00"})
        qty_fmt = workbook.add_format({"num_format": "#,##0.###"})

        for sheet_name, frame in sheets.items():
            _write_sheet(workbook, sheet_name, frame,
                         header_fmt, date_fmt, amount_fmt, qty_fmt)
    finally:
        workbook.close()

    log.info("  excel: %s (%.1f MB)", path, path.stat().st_size / 1048576)
    return path


def _write_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    frame: pd.DataFrame,
    header_fmt,
    date_fmt,
    amount_fmt,
    qty_fmt,
) -> None:
    sheet = workbook.add_worksheet(sheet_name[:31])
    columns = list(frame.columns)

    # Per-column format and width, decided once from the dtype.
    formats: list[object] = []
    for col in columns:
        dtype = frame[col].dtype
        if pd.api.types.is_datetime64_any_dtype(dtype):
            formats.append(date_fmt)
            width = 13
        elif col in _AMOUNT_COLUMNS:
            formats.append(amount_fmt)
            width = 14
        elif pd.api.types.is_bool_dtype(dtype):
            formats.append(None)
            width = 9
        elif pd.api.types.is_numeric_dtype(dtype):
            formats.append(qty_fmt)
            width = 11
        else:
            formats.append(None)
            width = min(max(len(col) + 2, 12), 40)
        sheet.set_column(len(formats) - 1, len(formats) - 1, width, formats[-1])

    for idx, col in enumerate(columns):
        sheet.write_string(0, idx, col, header_fmt)

    # itertuples is materially faster than iterrows and keeps native types.
    for row_ix, row in enumerate(frame.itertuples(index=False, name=None), start=1):
        for col_ix, value in enumerate(row):
            _write_cell(sheet, row_ix, col_ix, value, formats[col_ix])

    sheet.freeze_panes(1, 0)
    if len(frame):
        sheet.autofilter(0, 0, len(frame), max(len(columns) - 1, 0))
    log.info("  excel sheet %-14s %7d rows x %d cols", sheet_name, len(frame), len(columns))


def _write_cell(sheet, row: int, col: int, value, cell_format) -> None:
    """Write one cell, mapping every flavour of missing value to a blank."""
    if value is None or value is pd.NaT:
        sheet.write_blank(row, col, None, cell_format)
        return
    if isinstance(value, bool):
        sheet.write_boolean(row, col, value, cell_format)
        return
    if isinstance(value, (pd.Timestamp,)):
        sheet.write_datetime(row, col, value.to_pydatetime(), cell_format)
        return
    try:
        if pd.isna(value):
            sheet.write_blank(row, col, None, cell_format)
            return
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float)):
        sheet.write_number(row, col, value, cell_format)
        return
    sheet.write_string(row, col, str(value), cell_format)
