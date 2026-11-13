"""Extraction layer: BC entities and local reference tables -> typed DataFrames."""

from __future__ import annotations

import logging
from datetime import date, timedelta

import pandas as pd

import config
from bc_client import BCClient

log = logging.getLogger(__name__)

# Business Central represents an empty Date as 0001-01-01 rather than null.
# Anything at or before this cutoff is treated as "no date".
_BLANK_DATE_CUTOFF = pd.Timestamp("1900-01-01")


def coerce_dates(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Parse date columns, mapping BC's 0001-01-01 sentinel to NaT.

    This is a deliberate correction. The workbook converted the sentinel to a
    real year-1 date, so `ISBLANK(...)` never fired for a missing requested
    delivery date. Treating it as genuinely missing is what the formula
    intended -- see README, "Blank date handling".
    """
    for col in columns:
        if col not in frame.columns:
            continue
        parsed = pd.to_datetime(frame[col], errors="coerce", format="mixed")
        # Drop timezone if BC returned datetimes rather than plain dates.
        if isinstance(parsed.dtype, pd.DatetimeTZDtype):
            parsed = parsed.dt.tz_localize(None)
        frame[col] = parsed.mask(parsed <= _BLANK_DATE_CUTOFF).dt.normalize()
    return frame


def coerce_numbers(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    for col in columns:
        if col in frame.columns:
            frame[col] = pd.to_numeric(frame[col], errors="coerce")
    return frame


def _frame(records: list[dict], expected: list[str]) -> pd.DataFrame:
    """Build a DataFrame that always has the expected columns, even when empty."""
    frame = pd.DataFrame.from_records(records)
    for col in expected:
        if col not in frame.columns:
            frame[col] = pd.NA
    # Keep only expected columns, in a stable order.
    return frame[expected].copy()


def _text(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Normalise text keys: cast to str, strip, and blank out nulls."""
    for col in columns:
        if col in frame.columns:
            frame[col] = (
                frame[col].astype("string").fillna("").str.strip()
            )
    return frame


def _lookback_filter(field: str) -> str | None:
    """Build an OData date filter that still admits BC's blank-date sentinel."""
    if not config.LOOKBACK_DAYS:
        return None
    cutoff = date.today() - timedelta(days=config.LOOKBACK_DAYS)
    return f"({field} ge {cutoff.isoformat()} or {field} eq 0001-01-01)"


def _combine(*clauses: str | None) -> str | None:
    """AND together the non-empty filter clauses."""
    parts = [c for c in clauses if c]
    return " and ".join(parts) if parts else None


# --------------------------------------------------------------------------
# Order lines
# --------------------------------------------------------------------------
def archive_order_lines(client: BCClient) -> pd.DataFrame:
    """Archived sales order lines -- the spine of the OTIF report."""
    records = client.fetch(
        config.ENTITY_ARCHIVE_ORDER_LINES,
        select=config.SELECT_ARCHIVE_ORDER_LINES,
        odata_filter=_combine(
            config.FILTER_ARCHIVE_ORDER_LINES,
            _lookback_filter("Requested_Delivery_Date"),
        ),
    )
    frame = _frame(records, config.SELECT_ARCHIVE_ORDER_LINES)
    frame = _text(frame, ["Document_No", "Type", "No", "Description"])
    frame = coerce_numbers(frame, ["Version_No", "Line_No", "Quantity",
                                   "Line_Amount", "Qty_to_Ship"])
    frame = coerce_dates(frame, ["Requested_Delivery_Date"])

    before = len(frame)
    frame = frame[
        (frame["Version_No"] == 1)
        & (frame["Type"] == "Item")
        & (frame["No"] != "")
        & (~frame["No"].str.startswith(config.SCRAP_ITEM_PREFIX))
    ].copy()
    log.info("  archive lines: %d -> %d after Version_No=1 / Type=Item / SCRAP filters",
             before, len(frame))
    frame["Source"] = "Archive"
    return frame.drop(columns=["Version_No", "Type"])


def open_order_lines(client: BCClient) -> pd.DataFrame:
    """Live (not yet archived) sales order lines.

    Only used when INCLUDE_OPEN_ORDERS is true. Note there is no Version_No
    field on this entity -- that absence is exactly what silently dropped these
    rows in the Power Query version.
    """
    records = client.fetch(
        config.ENTITY_OPEN_ORDER_LINES,
        select=config.SELECT_OPEN_ORDER_LINES,
        odata_filter=_combine(
            config.FILTER_OPEN_ORDER_LINES,
            _lookback_filter("Requested_Delivery_Date"),
        ),
    )
    frame = _frame(records, config.SELECT_OPEN_ORDER_LINES)
    frame = _text(frame, ["Document_No", "Type", "No", "Description"])
    frame = coerce_numbers(frame, ["Line_No", "Quantity", "Line_Amount", "Qty_to_Ship"])
    frame = coerce_dates(frame, ["Requested_Delivery_Date"])

    before = len(frame)
    frame = frame[
        (frame["Type"] == "Item")
        & (frame["No"] != "")
        & (~frame["No"].str.startswith(config.SCRAP_ITEM_PREFIX))
    ].copy()
    log.info("  open lines: %d -> %d after Type=Item / SCRAP filters", before, len(frame))
    frame["Source"] = "Open"
    return frame.drop(columns=["Type"])


# --------------------------------------------------------------------------
# Shipments
# --------------------------------------------------------------------------
def shipment_lines(client: BCClient) -> pd.DataFrame:
    records = client.fetch(
        config.ENTITY_SHIPMENT_LINES,
        select=config.SELECT_SHIPMENT_LINES,
        odata_filter=config.FILTER_SHIPMENT_LINES,
    )
    frame = _frame(records, config.SELECT_SHIPMENT_LINES)
    frame = _text(frame, ["Document_No", "Type", "No", "Description",
                          "Description_2", "Location_Code",
                          "Unit_of_Measure_Code", "Unit_of_Measure"])
    frame = coerce_numbers(frame, ["Line_No", "Quantity", "Quantity_Invoiced"])

    before = len(frame)
    frame = frame[(frame["Quantity"] > 0) & (frame["Type"] == "Item")].copy()
    log.info("  shipment lines: %d -> %d after Quantity>0 / Type=Item", before, len(frame))
    return frame.drop(columns=["Type"])


def shipment_headers(client: BCClient) -> pd.DataFrame:
    records = client.fetch(
        config.ENTITY_SHIPMENT_HEADERS,
        select=config.SELECT_SHIPMENT_HEADERS,
    )
    frame = _frame(records, config.SELECT_SHIPMENT_HEADERS)
    frame = _text(frame, ["No", "Order_No"])
    frame = coerce_dates(frame, ["Posting_Date"])
    # VLOOKUP semantics: first match wins.
    return frame.drop_duplicates(subset=["No"], keep="first")


# --------------------------------------------------------------------------
# Order headers (live + archive, combined into one dimension)
# --------------------------------------------------------------------------
def _order_headers(client: BCClient, entity: str) -> pd.DataFrame:
    records = client.fetch(entity, select=config.SELECT_ORDER_HEADERS)
    frame = _frame(records, config.SELECT_ORDER_HEADERS)
    frame = _text(frame, ["No", "Sell_to_Customer_Name", "Business_SubVirtical",
                          "Ship_to_Post_Code", "Brand"])
    frame = coerce_dates(frame, ["PO_Date"])
    return frame.rename(columns={"PO_Date": "Order_Date"})


def order_headers(client: BCClient) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (combined_headers, archive_headers).

    ``combined_headers`` backs the Brand / Business Vertical / City lookups and
    matches the workbook's ``Sales_Orders`` table: live rows first, so they win
    on a duplicate order number (XLOOKUP returns the first match).

    ``archive_headers`` is returned separately because the workbook's PO Date
    column looks up the *archive* header table only.
    """
    live = _order_headers(client, config.ENTITY_ORDER_HEADERS_LIVE)
    archive = _order_headers(client, config.ENTITY_ORDER_HEADERS_ARCHIVE)

    archive = archive[archive["Brand"] != ""]
    for frame in (live, archive):
        frame.drop(
            frame.index[frame["Brand"].str.startswith(config.SCRAP_BRAND_PREFIX)],
            inplace=True,
        )

    combined = (
        pd.concat([live, archive], ignore_index=True)
        .drop_duplicates(subset=["No"], keep="first")
    )
    archive_pos = archive.drop_duplicates(subset=["No"], keep="first")
    log.info("  order headers: %d live + %d archive -> %d combined",
             len(live), len(archive), len(combined))
    return combined, archive_pos


# --------------------------------------------------------------------------
# Local reference tables (maintained by hand, exported from the workbook)
# --------------------------------------------------------------------------
def mta_list() -> pd.DataFrame:
    frame = pd.read_csv(config.MTA_LIST_CSV, dtype=str).fillna("")
    frame["Item_No"] = frame["Item_No"].str.strip()
    frame["MTA"] = (
        frame["MTA"].str.strip().str.upper().map({"TRUE": True, "FALSE": False, "1": True, "0": False})
    ).fillna(False)
    dupes = frame["Item_No"].duplicated().sum()
    if dupes:
        log.warning("  mta_list.csv has %d duplicate Item_No; keeping first", dupes)
        frame = frame.drop_duplicates(subset=["Item_No"], keep="first")
    log.info("  mta list: %d items (%d flagged MTA)", len(frame), int(frame["MTA"].sum()))
    return frame


def pincode_mapping() -> pd.DataFrame:
    frame = pd.read_csv(config.PINCODE_CSV, dtype=str).fillna("")
    for col in frame.columns:
        frame[col] = frame[col].str.strip()
    dupes = frame["Pincode"].duplicated().sum()
    if dupes:
        log.warning("  pincode_mapping.csv has %d duplicate Pincode; keeping first", dupes)
        frame = frame.drop_duplicates(subset=["Pincode"], keep="first")
    log.info("  pincode mapping: %d pincodes", len(frame))
    return frame
