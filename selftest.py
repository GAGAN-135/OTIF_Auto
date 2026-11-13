"""Offline verification of the OTIF logic against hand-computed cases.

Runs the real transform against small synthetic frames, so the business rules
can be checked without Business Central credentials.

    python selftest.py
"""

from __future__ import annotations

import logging
import sys
from datetime import date

import pandas as pd

import outputs
import transform

logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
log = logging.getLogger("selftest")

D = pd.Timestamp
PO = D("2026-01-01")


def _order_lines() -> pd.DataFrame:
    """One row per scenario. RDD = requested delivery date."""
    rows = [
        # doc,        line,  item,     qty,   rdd
        ("SO001",     10000, "ITEM-A", 100.0, D("2026-02-01")),  # 1 delivered in full, in time
        ("SO002",     10000, "ITEM-A",  50.0, D("2026-02-01")),  # 2 delivered late
        ("SO003",     10000, "ITEM-A", 100.0, D("2026-02-01")),  # 3 two part shipments = full
        ("SO004",     10000, "ITEM-M",  10.0, pd.NaT),           # 4 MTA item, blank RDD
        ("SO005",     10000, "ITEM-A", 100.0, D("2026-02-01")),  # 5 never delivered
        ("SO006",     10000, "ITEM-A",   5.0, D("2026-02-01")),  # 6 no archive header
        ("SO007",     10000, "ITEM-A", 100.0, D("2026-02-01")),  # 7 partial delivery
        ("SO1",           1, "ITEM-A",   7.0, D("2026-02-01")),  # 8a key-collision pair
        ("SO",           11, "ITEM-A",   7.0, D("2026-02-01")),  # 8b key-collision pair
        ("SCRAP001",  10000, "ITEM-A",  99.0, D("2026-02-01")),  # 9 scrap doc, dropped
        ("SO008",     10000, "ITEM-A",  20.0, D("2026-02-01")),  # 10 header exists, PO blank
    ]
    frame = pd.DataFrame(rows, columns=[
        "Document_No", "Line_No", "No", "Quantity", "Requested_Delivery_Date"])
    frame["Description"] = "Test item"
    frame["Line_Amount"] = frame["Quantity"] * 10
    frame["Qty_to_Ship"] = 0.0
    frame["Source"] = "Archive"
    for col in ("Document_No", "No", "Description", "Source"):
        frame[col] = frame[col].astype("string")
    return frame


def _shipments() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Shipment lines plus the headers that carry Order_No and Posting_Date."""
    # shipment doc, order,   posting date
    headers = [
        ("SH001", "SO001", D("2026-01-20")),   # inside SO001 deadline (28 Jan)
        ("SH002", "SO002", D("2026-02-10")),   # after SO002 deadline
        ("SH003", "SO003", D("2026-01-10")),
        ("SH004", "SO003", D("2026-01-20")),
        ("SH005", "SO004", D("2026-01-05")),   # inside MTA deadline (8 Jan)
        ("SH006", "SO006", D("2026-01-10")),   # order has no PO date
        ("SH007", "SO007", D("2026-01-10")),
        ("SH008", "SO1",   D("2026-01-10")),
        ("SH009", "SO",    D("2026-01-10")),
        ("SH010", "SO008", D("2026-01-15")),   # inside the RDD-4 fallback deadline
    ]
    # shipment doc, item, qty
    lines = [
        ("SH001", "ITEM-A", 100.0),
        ("SH002", "ITEM-A",  50.0),
        ("SH003", "ITEM-A",  60.0),
        ("SH004", "ITEM-A",  40.0),
        ("SH005", "ITEM-M",  10.0),
        ("SH006", "ITEM-A",   5.0),
        ("SH007", "ITEM-A",  60.0),   # partial: 60 of 100
        ("SH008", "ITEM-A",   7.0),
        ("SH009", "ITEM-A",   7.0),
        ("SH010", "ITEM-A",  20.0),
    ]
    hdr = pd.DataFrame(headers, columns=["No", "Order_No", "Posting_Date"])
    hdr["No"] = hdr["No"].astype("string")
    hdr["Order_No"] = hdr["Order_No"].astype("string")

    ln = pd.DataFrame(lines, columns=["Document_No", "No", "Quantity"])
    ln["Line_No"] = 10000
    ln["Description"] = "Test item"
    ln["Description_2"] = ""
    ln["Location_Code"] = "MAIN"
    ln["Unit_of_Measure_Code"] = "PCS"
    ln["Unit_of_Measure"] = "Piece"
    ln["Quantity_Invoiced"] = 0.0
    for col in ("Document_No", "No"):
        ln[col] = ln[col].astype("string")
    return ln, hdr


def _headers() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Order headers. SO006 is deliberately absent to test the missing-PO path."""
    rows = [
        ("SO001", PO), ("SO002", PO), ("SO003", PO), ("SO004", PO),
        ("SO005", PO), ("SO007", PO), ("SO1", PO), ("SO", PO),
        # Present, but PO_Date is BC's 0001-01-01 blank sentinel (-> NaT here).
        ("SO008", pd.NaT),
    ]
    frame = pd.DataFrame(rows, columns=["No", "Order_Date"])
    frame["Sell_to_Customer_Name"] = "Test Customer"
    frame["Business_SubVirtical"] = "Retail"
    frame["Brand"] = "BrandX"
    frame["Ship_to_Post_Code"] = "122004"
    frame["No"] = frame["No"].astype("string")
    frame["Ship_to_Post_Code"] = frame["Ship_to_Post_Code"].astype("string")
    return frame, frame[["No", "Order_Date"]].copy()


def _mta() -> pd.DataFrame:
    return pd.DataFrame({"Item_No": ["ITEM-A", "ITEM-M"], "MTA": [False, True]})


def _pincodes() -> pd.DataFrame:
    return pd.DataFrame({
        "Pincode": ["122004"], "City": ["Gurugram"],
        "State": ["Haryana"], "Country": ["India"],
    })


CHECKS: list[tuple[str, str, dict]] = [
    ("delivered in full before deadline", "SO001|ITEM-A", {
        "Final_Deadline": D("2026-01-28"), "Qty_Delivered_Before_Deadline": 100.0,
        "In_Full": True, "OTIF": True, "Delivery_Status": "Delivered"}),
    ("delivered but after the deadline", "SO002|ITEM-A", {
        "Final_Deadline": D("2026-01-28"), "Qty_Delivered_Before_Deadline": 0.0,
        "OTIF": False, "Delivery_Status": "Delivered"}),
    ("two part shipments sum to full", "SO003|ITEM-A", {
        "Qty_Delivered_Before_Deadline": 100.0, "OTIF": True}),
    ("MTA item uses PO+7 and blank RDD falls back to PO date", "SO004|ITEM-M", {
        "MTA": True, "Internal_Deadline": D("2026-01-08"),
        "Requested_Shipment_Date": PO, "Final_Deadline": D("2026-01-08"),
        "OTIF": True}),
    ("never delivered", "SO005|ITEM-A", {
        "Qty_Delivered_Before_Deadline": 0.0, "OTIF": False,
        "Delivery_Status": "Not delivered", "Posting_Date": pd.NaT}),
    ("missing archive header leaves no deadline", "SO006|ITEM-A", {
        "PO_Date": pd.NaT, "Final_Deadline": pd.NaT,
        "Qty_Delivered_Before_Deadline": 0.0, "OTIF": False}),
    # Distinct from the case above: the header exists but carries a blank PO
    # date, so the requested shipment date must still provide a deadline.
    ("blank PO date falls back to requested shipment date", "SO008|ITEM-A", {
        "PO_Date": pd.NaT, "Internal_Deadline": pd.NaT,
        "Requested_Shipment_Date": D("2026-01-28"),
        "Final_Deadline": D("2026-01-28"),
        "Qty_Delivered_Before_Deadline": 20.0, "OTIF": True}),
    ("partial delivery is not in full", "SO007|ITEM-A", {
        "Qty_Delivered_Before_Deadline": 60.0, "In_Full": False, "OTIF": False}),
    ("non-MTA item uses PO+14", "SO001|ITEM-A", {
        "Internal_Deadline": D("2026-01-15"), "MTA": False}),
]


def main() -> int:
    order_lines = _order_lines()
    ship_lines_raw, ship_headers = _shipments()
    combined_headers, archive_headers = _headers()

    combined_headers = transform.add_city(combined_headers, _pincodes())
    ship_lines = transform.prepare_shipment_lines(ship_lines_raw, ship_headers)
    otif = transform.build_otif_lines(
        order_lines=order_lines,
        ship_lines=ship_lines,
        archive_headers=archive_headers,
        combined_headers=combined_headers,
        mta=_mta(),
    )

    failures: list[str] = []
    indexed = otif.set_index("Key_Req")

    print()
    for label, key, expected in CHECKS:
        if key not in indexed.index:
            failures.append(f"{label}: key {key!r} missing from output")
            print(f"  FAIL  {label}  (key {key!r} not produced)")
            continue
        row = indexed.loc[key]
        bad = []
        for field, want in expected.items():
            got = row[field]
            ok = (pd.isna(got) and pd.isna(want)) or got == want
            if not ok:
                bad.append(f"{field}: expected {want!r}, got {got!r}")
        if bad:
            failures.extend(f"{label}: {b}" for b in bad)
            print(f"  FAIL  {label}")
            for b in bad:
                print(f"          {b}")
        else:
            print(f"  ok    {label}")

    # Structural checks -----------------------------------------------------
    print()
    scrap_rows = int(otif["Document_No"].str.startswith("SCRAP").sum())
    if scrap_rows:
        failures.append(f"scrap filter: {scrap_rows} SCRAP row(s) survived")
        print(f"  FAIL  scrap documents excluded ({scrap_rows} survived)")
    else:
        print("  ok    scrap documents excluded")

    # The delimiter fix: "SO1"+line 1 and "SO"+line 11 collided in the old key.
    collision_keys = {"SO1|1|ITEM-A|7", "SO|11|ITEM-A|7"}
    present = collision_keys & set(otif["Key"])
    if present != collision_keys:
        failures.append(f"key collision: only {present} survived, expected both")
        print(f"  FAIL  colliding keys kept separate (got {present})")
    else:
        print("  ok    colliding keys kept separate (old key merged these)")

    city_ok = otif.loc[otif["Document_No"] == "SO001", "City"].iloc[0] == "Gurugram"
    if not city_ok:
        failures.append("pincode -> city mapping did not resolve")
        print("  FAIL  pincode -> city mapping")
    else:
        print("  ok    pincode -> city mapping")

    # Deadline status labelling, as of a date sitting between the two clusters.
    as_of = D("2026-01-20")
    otif = transform.add_deadline_status(otif, as_of)
    labels = set(otif["Deadline_Status"])
    if not {"Elapsed", "Pending"} <= labels:
        failures.append(f"deadline status labels incomplete: {labels}")
        print(f"  FAIL  deadline status labelling (got {labels})")
    else:
        print(f"  ok    deadline status labelling ({sorted(labels)})")

    mislabelled = otif[
        (otif["Deadline_Status"] == "Pending") & (otif["Final_Deadline"] <= as_of)
    ]
    if len(mislabelled):
        failures.append(f"{len(mislabelled)} line(s) labelled Pending with an elapsed deadline")
        print("  FAIL  Pending label matches a future deadline")
    else:
        print("  ok    Pending label matches a future deadline")

    # Qty_Short must never go negative, even when over-delivered.
    if (otif["Qty_Short"] < 0).any():
        failures.append("Qty_Short went negative")
        print("  FAIL  Qty_Short is never negative")
    else:
        print("  ok    Qty_Short is never negative")

    pending_view = transform.build_pending(otif)
    expected_pending = int((otif["Deadline_Status"] == "Pending").sum())
    if len(pending_view) != expected_pending:
        failures.append(
            f"pending view has {len(pending_view)} rows, expected {expected_pending}")
        print("  FAIL  pending view row count")
    elif not pending_view["Final_Deadline"].is_monotonic_increasing:
        failures.append("pending view is not sorted by deadline")
        print("  FAIL  pending view sorted soonest-first")
    else:
        print(f"  ok    pending view ({len(pending_view)} row(s), soonest first)")

    snop = transform.build_snop(otif)
    if len(snop) != otif["Key_Req"].nunique():
        failures.append("snop grain is not one row per order-item")
        print("  FAIL  SNOP is one row per order-item")
    else:
        print("  ok    SNOP is one row per order-item")

    # OTIF is not assessable before the deadline arrives. With a cutoff between
    # the two deadline clusters, only the elapsed ones should survive.
    cutoff = D("2026-01-20")
    gated = transform.build_snop(otif, deadline_cutoff=cutoff)
    late = gated["Final_Deadline"].dropna()
    if len(late) and late.max() > cutoff:
        failures.append("snop cutoff kept a deadline after the cutoff date")
        print(f"  FAIL  SNOP deadline cutoff (max kept = {late.max()})")
    elif len(gated) >= len(snop):
        failures.append("snop cutoff excluded nothing despite future deadlines")
        print("  FAIL  SNOP deadline cutoff excluded nothing")
    else:
        print(f"  ok    SNOP deadline cutoff ({len(snop)} -> {len(gated)} rows)")

    # A line with no computable deadline must survive the cutoff, so a missing
    # archive header stays visible rather than silently disappearing.
    if "SO006|ITEM-A" not in set(gated["Key_Req"]):
        failures.append("snop cutoff dropped the null-deadline line")
        print("  FAIL  SNOP cutoff keeps null-deadline lines")
    else:
        print("  ok    SNOP cutoff keeps null-deadline lines")

    # Exercise both writers end to end so format bugs surface here, not at 6am.
    print()
    tmp = transform.select_columns(otif, transform.OTIF_COLUMNS)
    snop_out = transform.select_columns(snop, transform.SNOP_COLUMNS)
    pending_out = transform.select_columns(pending_view, transform.PENDING_COLUMNS)
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            outputs.write_excel({"OTIF Lines": tmp, "SNOP": snop_out,
                                 "Pending": pending_out},
                                Path(tmpdir) / "selftest.xlsx")
            print("  ok    Excel writer")
        except Exception as exc:  # noqa: BLE001 - report and keep going
            failures.append(f"excel writer: {exc}")
            print(f"  FAIL  Excel writer: {exc}")
        try:
            outputs.write_parquet({"otif_lines": tmp, "snop": snop_out},
                                  run_date=date(2026, 8, 4),
                                  parquet_dir=Path(tmpdir), keep_history=False)
            print("  ok    Parquet writer")
        except Exception as exc:  # noqa: BLE001
            failures.append(f"parquet writer: {exc}")
            print(f"  FAIL  Parquet writer: {exc}")

    print()
    print(f"{len(CHECKS) + 4} scenario group(s) checked, {len(failures)} failure(s)")
    if failures:
        print("\nFailures:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
