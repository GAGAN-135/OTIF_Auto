"""The OTIF calculation.

This module is the Python equivalent of the workbook's 1,157,976 formula cells.
Every column below maps to a named column in the old `SaleOrderLineOTIF` sheet;
the mapping is spelled out in README under "Formula translation table".

Where the spreadsheet did an XLOOKUP (a linear scan per row) this does a hash
join, which is why the whole thing finishes in seconds.
"""

from __future__ import annotations

import logging

import numpy as np
import pandas as pd

import config

log = logging.getLogger(__name__)

# Composite keys are joined with this separator. The workbook concatenated the
# parts with nothing between them, so document "SO1" + line "1" produced the
# same key as "SO" + line "11" -- and because Key is the de-duplication key, a
# collision silently deleted a real order line.
KEY_SEP = "|"


def _key(*parts: pd.Series) -> pd.Series:
    """Build a delimited composite key from several columns."""
    as_text = [p.astype("string").fillna("") for p in parts]
    out = as_text[0]
    for part in as_text[1:]:
        out = out + KEY_SEP + part
    return out


def _quantity_text(series: pd.Series) -> pd.Series:
    """Render quantity for the key the way `Text.From` did: 10.0 -> "10"."""
    def fmt(value: float) -> str:
        if pd.isna(value):
            return ""
        if float(value).is_integer():
            return str(int(value))
        return repr(float(value))

    return series.map(fmt).astype("string")


# --------------------------------------------------------------------------
# Shipment lines: attach the sale order number and posting date from the header
# --------------------------------------------------------------------------
def prepare_shipment_lines(
    ship_lines: pd.DataFrame, ship_headers: pd.DataFrame
) -> pd.DataFrame:
    """Adds `Sale_Order`, `Posting_Date`, and `Key` (was columns L, N, M)."""
    frame = ship_lines.copy()
    header_map = ship_headers.set_index("No")
    frame["Sale_Order"] = frame["Document_No"].map(header_map["Order_No"]).fillna("")
    frame["Posting_Date"] = frame["Document_No"].map(header_map["Posting_Date"])
    frame["Key"] = _key(frame["Sale_Order"], frame["No"])

    orphans = int((frame["Sale_Order"] == "").sum())
    if orphans:
        log.warning("  %d shipment line(s) have no matching shipment header", orphans)
    return frame


# --------------------------------------------------------------------------
# The OTIF spine
# --------------------------------------------------------------------------
def build_otif_lines(
    order_lines: pd.DataFrame,
    ship_lines: pd.DataFrame,
    archive_headers: pd.DataFrame,
    combined_headers: pd.DataFrame,
    mta: pd.DataFrame,
) -> pd.DataFrame:
    """Return the OTIF line-level table (old sheet: SaleOrderLineOTIF, cols A-W)."""
    frame = order_lines.copy()

    # Drop documents flagged as scrap (M: two separate SCRAP/SCR filters).
    before = len(frame)
    frame = frame[
        ~frame["Document_No"].str.upper().str.startswith(config.SCRAP_DOC_PREFIXES)
    ].copy()
    if before != len(frame):
        log.info("  dropped %d scrap document line(s)", before - len(frame))

    # --- One row per physical order line (was implicit: archive was the only
    # source) --------------------------------------------------------------
    # With open orders enabled a line appears twice: once live while the order
    # is open, once archived after it closes. They are the same physical line,
    # so only one may survive, and the archived version wins because it is the
    # final, closed state.
    #
    # This cannot be left to the Key de-duplication below. Key includes the
    # quantity, so a line whose quantity was edited before the order closed
    # produces two different Keys and would be counted twice.
    #
    # The item number must be part of the identity: BC reuses a line number
    # within a document (SO-M12803 line 10000 carries two different items), so
    # de-duplicating on document+line alone would delete a real order line.
    # Document+line+item is unique within each source.
    if "Source" in frame.columns and frame["Source"].nunique() > 1:
        before = len(frame)
        frame = (
            frame.assign(_archive_first=(frame["Source"] != "Archive").astype(int))
            .sort_values("_archive_first", kind="stable")
            .drop_duplicates(subset=["Document_No", "Line_No", "No"], keep="first")
            .drop(columns="_archive_first")
            .reset_index(drop=True)
        )
        if before != len(frame):
            log.info("  collapsed %d open line(s) superseded by their archived "
                     "version", before - len(frame))

    # --- Key / de-duplication (was col I) ---------------------------------
    frame["Key"] = _key(
        frame["Document_No"],
        frame["Line_No"].astype("Int64").astype("string"),
        frame["No"],
        _quantity_text(frame["Quantity"]),
    )
    before = len(frame)
    frame = frame.drop_duplicates(subset=["Key"], keep="first").reset_index(drop=True)
    if before != len(frame):
        log.info("  removed %d duplicate line(s) on Key", before - len(frame))

    # --- Key for requested qty (was col J) --------------------------------
    frame["Key_Req"] = _key(frame["Document_No"], frame["No"])

    # --- PO Date (was col K) ----------------------------------------------
    # The ARCHIVE header is the primary source, exactly as the workbook did, so
    # every already-archived line keeps the PO Date it has always had and no
    # historical figure is restated by this lookup.
    #
    # A live (not yet archived) order has no archive header, so it falls back to
    # the live order header's Order_Date. Without that fallback an open line
    # gets no PO Date, therefore no Final Deadline, therefore no deadline month
    # -- it drops out of every month bucket until the order closes, and then
    # reappears inside an already-published month. That is what made a closed
    # month keep moving; see README "A closed month must not move".
    archive_map = archive_headers.set_index("No")["Order_Date"]
    live_map = (
        combined_headers.drop_duplicates(subset=["No"]).set_index("No")["Order_Date"]
    )

    in_archive = frame["Document_No"].isin(archive_map.index)
    frame["PO_Date"] = frame["Document_No"].map(archive_map).where(
        in_archive, frame["Document_No"].map(live_map)
    )
    from_live = ~in_archive & frame["Document_No"].isin(live_map.index)
    if from_live.any():
        log.info("  %d line(s) are not archived yet -> PO Date taken from the "
                 "live order header", int(from_live.sum()))

    # Two different situations must be told apart, because the workbook treated
    # them differently and only one of them means "no deadline":
    #   * document absent from BOTH header tables -> XLOOKUP returned #N/A,
    #     which propagated through MAX into IFNA(...,0). No usable deadline.
    #   * document present but PO_Date is BC's 0001-01-01 blank sentinel -> the
    #     workbook's MAX simply fell back to the requested shipment date.
    header_found = in_archive | from_live
    blank_po = header_found & frame["PO_Date"].isna()
    if (~header_found).any():
        log.warning("  %d line(s) (%.1f%%) have no order header at all "
                    "-> no deadline, so OTIF = False",
                    int((~header_found).sum()),
                    100.0 * int((~header_found).sum()) / max(len(frame), 1))
    if blank_po.any():
        log.info("  %d line(s) have an archive header with a blank PO Date "
                 "-> deadline falls back to the requested shipment date",
                 int(blank_po.sum()))

    # --- MTA (was col L) --------------------------------------------------
    mta_map = mta.set_index("Item_No")["MTA"]
    frame["MTA"] = frame["No"].map(mta_map).fillna(False).astype(bool)

    # --- Internal Deadline (was col M) ------------------------------------
    lead_days = np.where(frame["MTA"], config.LEAD_DAYS_MTA, config.LEAD_DAYS_NON_MTA)
    frame["Internal_Deadline"] = frame["PO_Date"] + pd.to_timedelta(lead_days, unit="D")

    # --- Requested Shipment date (was col N) ------------------------------
    shifted = frame["Requested_Delivery_Date"] - pd.Timedelta(
        days=config.SHIPMENT_LEAD_OFFSET_DAYS
    )
    frame["Requested_Shipment_Date"] = shifted.where(
        frame["Requested_Delivery_Date"].notna(), frame["PO_Date"]
    )

    # --- Final Deadline (was col O) ---------------------------------------
    # MAX over the two deadlines, skipping whichever could not be computed --
    # the same thing Excel's MAX did. Only a document missing from both header
    # tables yields no deadline (the #N/A -> IFNA(...,0) path); a blank PO Date
    # still leaves the requested shipment date to work with.
    frame["Final_Deadline"] = frame[
        ["Internal_Deadline", "Requested_Shipment_Date"]
    ].max(axis=1)
    frame.loc[~header_found, "Final_Deadline"] = pd.NaT

    # --- Qty Delivered before deadline (was col P, the SUMIFS) ------------
    frame["Qty_Delivered_Before_Deadline"] = _qty_before_deadline(frame, ship_lines)

    # --- Posting Date / delivery status (was col Q) -----------------------
    # The workbook put either a date or the text "Not delivered" in one column.
    # A mixed-type column breaks Parquet and confuses Power BI, so it is split
    # into a real date plus a status flag.
    first_posting = (
        ship_lines.dropna(subset=["Posting_Date"])
        .drop_duplicates(subset=["Key"], keep="first")
        .set_index("Key")["Posting_Date"]
    )
    frame["Posting_Date"] = frame["Key_Req"].map(first_posting)
    frame["Delivery_Status"] = np.where(
        frame["Posting_Date"].notna(), "Delivered", config.NOT_DELIVERED_LABEL
    )

    # --- In Full / On time / OTIF (was cols R, S, T) ----------------------
    # Note: the workbook defined all three identically (`On time` and `OTIF`
    # were literal copies of `In Full`). Timeliness is already enforced by the
    # deadline filter inside the quantity sum above, so the OTIF number is
    # sound -- but `On_Time` measures nothing independently. Kept as-is for
    # parity; see README "On time is not an independent measure".
    frame["In_Full"] = (
        frame["Qty_Delivered_Before_Deadline"] >= frame["Quantity"].fillna(0)
    )
    frame["On_Time"] = frame["In_Full"]
    frame["OTIF"] = frame["In_Full"]

    # --- Business Vertical / Brand / City (was cols U, V, W) --------------
    headers = combined_headers.set_index("No")
    frame["Business_Vertical"] = frame["Document_No"].map(headers["Business_SubVirtical"])
    frame["Brand"] = frame["Document_No"].map(headers["Brand"])
    frame["City"] = frame["Document_No"].map(headers["City"])
    frame["Customer_Name"] = frame["Document_No"].map(headers["Sell_to_Customer_Name"])

    return frame


def _qty_before_deadline(lines: pd.DataFrame, ship_lines: pd.DataFrame) -> pd.Series:
    """Vectorised replacement for the per-row SUMIFS.

    SUMIFS(ship[Quantity], ship[Key], Key_Req, ship[Posting Date], "<=" Final_Deadline)
    """
    zero = pd.Series(0.0, index=lines.index)
    if ship_lines.empty or lines.empty:
        return zero

    left = (
        lines[["Key_Req", "Final_Deadline"]]
        .reset_index()
        .rename(columns={"index": "_line_ix"})
    )
    right = ship_lines[["Key", "Quantity", "Posting_Date"]]
    joined = left.merge(right, left_on="Key_Req", right_on="Key", how="inner")

    within = (
        joined["Posting_Date"].notna()
        & joined["Final_Deadline"].notna()
        & (joined["Posting_Date"] <= joined["Final_Deadline"])
    )
    sums = joined.loc[within].groupby("_line_ix")["Quantity"].sum()
    return sums.reindex(lines.index).fillna(0.0)


# --------------------------------------------------------------------------
# Dimension: sales orders (adds City from the pincode map)
# --------------------------------------------------------------------------
def add_city(headers: pd.DataFrame, pincodes: pd.DataFrame) -> pd.DataFrame:
    """Adds `City` by pincode (old `SO PO Date` sheet, col G)."""
    frame = headers.copy()
    pin_map = pincodes.set_index("Pincode")
    key = frame["Ship_to_Post_Code"].astype("string").fillna("").str.strip()
    frame["City"] = key.map(pin_map["City"])
    frame["State"] = key.map(pin_map["State"])

    unmapped = frame.loc[frame["City"].isna() & (key != ""), "Ship_to_Post_Code"]
    if len(unmapped):
        log.warning("  %d order(s) have a pincode missing from pincode_mapping.csv "
                    "(%d distinct, e.g. %s)",
                    len(unmapped), unmapped.nunique(),
                    ", ".join(sorted(unmapped.unique())[:5]))
    return frame


# --------------------------------------------------------------------------
# Deadline status
# --------------------------------------------------------------------------
DEADLINE_ELAPSED = "Elapsed"
DEADLINE_PENDING = "Pending"
DEADLINE_NONE = "No deadline"


def add_deadline_status(frame: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    """Label each line by whether its deadline has arrived, and by how much it is short.

    Makes the not-yet-due population directly filterable instead of requiring a
    date comparison downstream, and keeps that definition identical everywhere.

    * ``Elapsed``     - deadline has passed; OTIF is meaningful.
    * ``Pending``     - deadline is in the future; OTIF is not yet decided.
    * ``No deadline`` - neither a PO date nor a requested delivery date exists,
                       so no deadline can be derived at all.
    """
    deadline = frame["Final_Deadline"]
    frame["Deadline_Status"] = np.select(
        [deadline.isna(), deadline > as_of],
        [DEADLINE_NONE, DEADLINE_PENDING],
        default=DEADLINE_ELAPSED,
    )
    # Negative = the deadline is already behind us.
    frame["Days_To_Deadline"] = (deadline - as_of).dt.days.astype("Int64")
    frame["Qty_Short"] = (
        frame["Quantity"].fillna(0) - frame["Qty_Delivered_Before_Deadline"].fillna(0)
    ).clip(lower=0)
    return frame


# --------------------------------------------------------------------------
# SNOP view: one row per order + item
# --------------------------------------------------------------------------
def build_snop(
    otif_lines: pd.DataFrame,
    deadline_cutoff: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """Collapse OTIF lines to order-item grain (old `SNOP` sheet).

    ``deadline_cutoff`` restricts the result to lines whose Final Deadline has
    actually arrived, which is what makes month-by-month OTIF meaningful: a line
    that is undelivered but not yet due is pending, not failed, and counting it
    as a failure understates OTIF. Pass None to keep every line.

    Lines with no computable deadline are kept. The workbook's `IFNA(...,0)` put
    them at date zero, so they always passed its cutoff, and keeping them means a
    missing archive header stays visible instead of silently vanishing.
    """
    if otif_lines.empty:
        return otif_lines.copy()

    snop = otif_lines.drop_duplicates(subset=["Key_Req"], keep="first").copy()

    if deadline_cutoff is not None:
        before = len(snop)
        elapsed = snop["Final_Deadline"].isna() | (snop["Final_Deadline"] <= deadline_cutoff)
        snop = snop[elapsed].copy()
        log.info("  snop: excluded %d order-item(s) whose deadline is after %s "
                 "(not yet assessable)",
                 before - len(snop), deadline_cutoff.date())

    # The workbook keyed SNOP's PO Date and Req delivery date on Document_No
    # rather than on the order-item key, so both show the first line's value for
    # the whole document. Reproduced here for parity.
    by_doc = otif_lines.drop_duplicates(subset=["Document_No"], keep="first").set_index("Document_No")
    snop["PO_Date"] = snop["Document_No"].map(by_doc["PO_Date"])
    snop["Requested_Delivery_Date"] = snop["Document_No"].map(
        by_doc["Requested_Delivery_Date"]
    )
    log.info("  snop: %d order-item rows from %d order lines", len(snop), len(otif_lines))
    return snop


# --------------------------------------------------------------------------
# Output column order
# --------------------------------------------------------------------------
OTIF_COLUMNS = [
    "Document_No", "Line_No", "No", "Description", "Quantity", "Line_Amount",
    "Qty_to_Ship", "Requested_Delivery_Date", "Key", "Key_Req", "PO_Date",
    "MTA", "Internal_Deadline", "Requested_Shipment_Date", "Final_Deadline",
    "Deadline_Status", "Days_To_Deadline",
    "Qty_Delivered_Before_Deadline", "Qty_Short", "Posting_Date",
    "Delivery_Status", "In_Full", "On_Time", "OTIF",
    "Business_Vertical", "Brand", "City", "Customer_Name", "Source",
]

SNOP_COLUMNS = [
    "Key_Req", "Document_No", "Customer_Name", "Brand", "Business_Vertical",
    "No", "Description", "PO_Date", "Requested_Delivery_Date", "MTA",
    "Internal_Deadline", "Requested_Shipment_Date", "Final_Deadline",
    "Deadline_Status", "Qty_Delivered_Before_Deadline", "Qty_Short",
    "Posting_Date", "Delivery_Status", "In_Full", "On_Time", "OTIF",
    "Line_Amount", "City", "Source",
]

# Forward order book: lines whose deadline has not arrived yet. Ordered so the
# most urgent shortfalls read off the top.
PENDING_COLUMNS = [
    "Final_Deadline", "Days_To_Deadline", "Document_No", "Line_No", "No",
    "Description", "Customer_Name", "Brand", "Business_Vertical", "City",
    "Quantity", "Qty_Delivered_Before_Deadline", "Qty_Short",
    "Delivery_Status", "In_Full", "MTA", "PO_Date",
    "Requested_Delivery_Date", "Line_Amount", "Key_Req", "Source",
]


def build_pending(otif_lines: pd.DataFrame) -> pd.DataFrame:
    """Lines whose deadline is still in the future, soonest and shortest first."""
    if otif_lines.empty or "Deadline_Status" not in otif_lines.columns:
        return otif_lines.iloc[0:0].copy()
    pending = otif_lines[otif_lines["Deadline_Status"] == DEADLINE_PENDING].copy()
    return pending.sort_values(
        ["Final_Deadline", "Qty_Short", "Document_No", "Line_No"],
        ascending=[True, False, True, True],
    )


def select_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Reorder to the published column list, tolerating absent columns."""
    present = [c for c in columns if c in frame.columns]
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        log.warning("  output is missing expected column(s): %s", ", ".join(missing))
    return frame[present].copy()
