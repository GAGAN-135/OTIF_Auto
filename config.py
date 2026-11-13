"""Central configuration for the OTIF pipeline.

Everything environment-specific lives in .env (see .env.example). Everything
business-rule-specific lives here as a named constant, so the rules are
auditable in one place instead of buried in 1.16 million spreadsheet formulas.
"""

from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import quote

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent
load_dotenv(PROJECT_ROOT / ".env")


def _flag(name: str, default: bool = False) -> bool:
    raw = os.getenv(name, "").strip().lower()
    if not raw:
        return default
    return raw in {"1", "true", "yes", "y", "on"}


def _int_or_none(name: str) -> int | None:
    raw = os.getenv(name, "").strip()
    return int(raw) if raw else None


# --------------------------------------------------------------------------
# Connection
# --------------------------------------------------------------------------
BC_USERNAME = os.getenv("BC_USERNAME", "").strip()
BC_ACCESS_KEY = os.getenv("BC_ACCESS_KEY", "").strip()
BC_BASE_URL = os.getenv("BC_BASE_URL", "http://20.235.106.184:8058/ERP/ODataV4").strip().rstrip("/")
BC_COMPANY = os.getenv("BC_COMPANY", "NUTASTE FOOD & DRINK LABS (P)").strip()

REQUEST_TIMEOUT = int(os.getenv("BC_TIMEOUT", "180"))  # seconds per HTTP request
# Note: we deliberately never send $top. In OData $top caps the *whole* result
# set, not the page size -- sending it would silently truncate. Server-side
# paging is driven by @odata.nextLink, which the client follows to exhaustion.


def company_url() -> str:
    """OData company root, with the company name percent-encoded."""
    return f"{BC_BASE_URL}/Company('{quote(BC_COMPANY, safe='')}')"


# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR") or PROJECT_ROOT / "output")
PARQUET_DIR = Path(os.getenv("PARQUET_DIR") or OUTPUT_DIR / "parquet")
HISTORY_DIR = PARQUET_DIR / "history"
LOG_DIR = PROJECT_ROOT / "logs"
REFERENCE_DIR = PROJECT_ROOT / "reference"

# Output names for this (PO-date) report. The "PO" prefix tells its files apart
# from release_otif.py's "release_" / "(Release Date)" files in the same folders.
EXCEL_NAME = "PO OTIF Report.xlsx"
PARQUET_PREFIX = "PO_"
MTA_LIST_CSV = REFERENCE_DIR / "mta_list.csv"
PINCODE_CSV = REFERENCE_DIR / "pincode_mapping.csv"

# --------------------------------------------------------------------------
# Behaviour switches
# --------------------------------------------------------------------------
# Live, not-yet-archived order lines. Defaults ON: with it off, a line only
# enters the report once BC archives its order (i.e. when the order closes),
# which lets an already-published month keep changing weeks later. See README
# "A closed month must not move".
INCLUDE_OPEN_ORDERS = _flag("INCLUDE_OPEN_ORDERS", True)
KEEP_HISTORY = _flag("KEEP_HISTORY", True)
LOOKBACK_DAYS = _int_or_none("LOOKBACK_DAYS")

# Escape hatch for a broken/slow source. ArchiveSalesOrdeLinesBi currently times
# out server-side even for a single row, so this allows a run from live order
# lines alone. Note the *archive header* entity is separate and healthy, so PO
# Date lookups still work. See README "When the archive endpoint times out".
SKIP_ARCHIVE_LINES = _flag("SKIP_ARCHIVE_LINES", False)

# --------------------------------------------------------------------------
# Business rules (transcribed from the workbook's formula columns)
# --------------------------------------------------------------------------
# Internal Deadline = PO Date + N days, where N depends on the item's MTA flag.
LEAD_DAYS_MTA = 7        # MTA = TRUE  -> PO Date + 7
LEAD_DAYS_NON_MTA = 14   # MTA = FALSE -> PO Date + 14

# Requested Shipment date = Requested_Delivery_Date - this many days,
# falling back to PO Date when the requested delivery date is blank.
SHIPMENT_LEAD_OFFSET_DAYS = 4

# Rows whose item number or document number starts with these are excluded.
SCRAP_ITEM_PREFIX = "SCRAP"
SCRAP_DOC_PREFIXES = ("SCRAP", "SCR")
SCRAP_BRAND_PREFIX = "SCRAP"

NOT_DELIVERED_LABEL = "Not delivered"

# OTIF cannot be judged before the deadline arrives: an undelivered line whose
# deadline is still in the future is not late, it is pending. Including such
# lines scores them as failures and understates OTIF. The workbook handled this
# by filtering SNOP to elapsed deadlines only (verified: SNOP contained no line
# with a deadline after the file's save date). SNOP is therefore the
# "assessable" view; otif_lines stays complete, future deadlines included.
SNOP_REQUIRE_ELAPSED_DEADLINE = _flag("SNOP_REQUIRE_ELAPSED_DEADLINE", True)

# --------------------------------------------------------------------------
# OData projections
#
# Each list is the *minimum* set of fields the transform actually consumes.
# The original Power Query pulled all 115 columns of SaleOrdersLineBi and then
# discarded ~100 of them client-side; these projections are why the Python run
# transfers a fraction of the data.
# --------------------------------------------------------------------------
SELECT_ARCHIVE_ORDER_LINES = [
    "Document_No", "Version_No", "Line_No", "Type", "No", "Description",
    "Quantity", "Line_Amount", "Qty_to_Ship", "Requested_Delivery_Date",
]

SELECT_OPEN_ORDER_LINES = [
    "Document_No", "Line_No", "Type", "No", "Description",
    "Quantity", "Line_Amount", "Qty_to_Ship", "Requested_Delivery_Date",
]

SELECT_SHIPMENT_LINES = [
    "Document_No", "Line_No", "Type", "No", "Description", "Description_2",
    "Location_Code", "Quantity", "Unit_of_Measure_Code", "Unit_of_Measure",
    "Quantity_Invoiced",
]

SELECT_SHIPMENT_HEADERS = ["No", "Order_No", "Posting_Date"]

SELECT_ORDER_HEADERS = [
    "No", "PO_Date", "Sell_to_Customer_Name", "Business_SubVirtical",
    "Ship_to_Post_Code", "Brand",
]

# Server-side row filters. Verified accepted by this service via $count, with
# large payload reductions (archive 102,773 -> 40,110; shipment lines
# 92,241 -> 37,782). The equivalent Python-side filters in extract.py are kept
# as a safety net in case a filter is ever rejected and has to be removed here.
FILTER_ARCHIVE_ORDER_LINES = "Version_No eq 1 and Type eq 'Item' and No ne ''"
FILTER_OPEN_ORDER_LINES = "Type eq 'Item'"
FILTER_SHIPMENT_LINES = "Type eq 'Item' and Quantity gt 0"

# Entity set names on the OData service.
ENTITY_ARCHIVE_ORDER_LINES = "ArchiveSalesOrdeLinesBi"
ENTITY_OPEN_ORDER_LINES = "SaleOrdersLineBi"
ENTITY_SHIPMENT_LINES = "PostedSalesShipmentLinesBi"
ENTITY_SHIPMENT_HEADERS = "PostedSalesShipmentBi"
ENTITY_ORDER_HEADERS_LIVE = "SalesOrder"
ENTITY_ORDER_HEADERS_ARCHIVE = "SalesOrderArchivesHeaderBi"


def validate() -> list[str]:
    """Return a list of configuration problems; empty means good to go."""
    problems: list[str] = []
    if not BC_USERNAME:
        problems.append("BC_USERNAME is not set (copy .env.example to .env).")
    if not BC_ACCESS_KEY:
        problems.append("BC_ACCESS_KEY is not set (Web Service Access Key from BC).")
    if not BC_BASE_URL.startswith(("http://", "https://")):
        problems.append(f"BC_BASE_URL looks malformed: {BC_BASE_URL!r}")
    for csv_path, label in ((MTA_LIST_CSV, "MTA list"), (PINCODE_CSV, "Pincode mapping")):
        if not csv_path.exists():
            problems.append(f"{label} reference file missing: {csv_path}")
    return problems
