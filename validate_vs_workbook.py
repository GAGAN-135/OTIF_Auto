"""Reconcile the Python output against the old workbook's cached values.

Reads the SaleOrderLineOTIF sheet straight out of the .xlsx (no Excel needed,
no recalculation) and compares it to output/parquet/current/PO_otif_lines.parquet
at line grain.

    python validate_vs_workbook.py "C:\\path\\to\\OTIF - New Copy.xlsx"
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import unescape

import pandas as pd

import config

# SaleOrderLineOTIF column letters -> meaning
COLS = {
    "A": "Document_No", "B": "Line_No", "C": "No", "E": "Quantity",
    "L": "MTA", "O": "Final_Deadline", "P": "Qty_Delivered_Before_Deadline",
    "T": "OTIF",
}
SHEET = "xl/worksheets/sheet8.xml"
EXCEL_EPOCH = pd.Timestamp("1899-12-30")

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)
ROW = re.compile(r"<row [^>]*?r=\"(\d+)\"[^>]*>(.*?)</row>", re.S)


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    return [
        unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S)))
        for si in re.findall(r"<si>(.*?)</si>", data, re.S)
    ]


def read_workbook(path: Path) -> pd.DataFrame:
    with zipfile.ZipFile(path) as zf:
        ss = _shared_strings(zf)
        xml = zf.read(SHEET).decode("utf-8")

    records: list[dict] = []
    for m in ROW.finditer(xml):
        rn, body = int(m.group(1)), m.group(2)
        if rn == 1:
            continue
        row: dict = {}
        for cm in CELL.finditer(body):
            col, _r, attrs, cbody = cm.groups()
            field = COLS.get(col)
            if not field:
                continue
            v = re.search(r"<v>(.*?)</v>", cbody, re.S)
            if not v:
                continue
            raw = unescape(v.group(1))
            if 't="s"' in attrs:
                idx = int(raw)
                row[field] = ss[idx] if idx < len(ss) else ""
            elif 't="b"' in attrs:
                row[field] = raw == "1"
            elif 't="str"' in attrs or 't="e"' in attrs:
                row[field] = raw
            else:
                row[field] = raw
        if row.get("Document_No"):
            records.append(row)

    frame = pd.DataFrame.from_records(records)
    for col in ("Line_No", "Quantity", "Qty_Delivered_Before_Deadline"):
        frame[col] = pd.to_numeric(frame[col], errors="coerce")
    # Final Deadline is an Excel date serial; 0 was the IFNA fallback.
    serial = pd.to_numeric(frame["Final_Deadline"], errors="coerce")
    frame["Final_Deadline"] = (
        EXCEL_EPOCH + pd.to_timedelta(serial.where(serial > 0), unit="D")
    )
    for col in ("MTA", "OTIF"):
        frame[col] = frame[col].map(lambda x: x is True)
    return frame


def line_key(frame: pd.DataFrame) -> pd.Series:
    return (
        frame["Document_No"].astype("string").fillna("")
        + "|" + pd.to_numeric(frame["Line_No"], errors="coerce").astype("Int64").astype("string")
        + "|" + frame["No"].astype("string").fillna("")
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="path to the old .xlsx")
    parser.add_argument("--parquet", type=Path,
                        default=config.PARQUET_DIR / "current"
                        / f"{config.PARQUET_PREFIX}otif_lines.parquet")
    args = parser.parse_args(argv)

    if not args.workbook.exists():
        print(f"workbook not found: {args.workbook}")
        return 1
    if not args.parquet.exists():
        print(f"parquet not found: {args.parquet} -- run main.py first")
        return 1

    print("Reading the old workbook's cached values...")
    old = read_workbook(args.workbook)
    new = pd.read_parquet(args.parquet)
    old["_k"] = line_key(old)
    new["_k"] = line_key(new)

    print()
    print(f"old workbook rows : {len(old):>7,}   distinct line keys: {old['_k'].nunique():>7,}")
    print(f"new pipeline rows : {len(new):>7,}   distinct line keys: {new['_k'].nunique():>7,}")

    old_otif = int(old["OTIF"].sum())
    new_otif = int(new["OTIF"].sum())
    print()
    print(f"old OTIF true     : {old_otif:>7,}  ({100.0*old_otif/max(len(old),1):5.1f}%)")
    print(f"new OTIF true     : {new_otif:>7,}  ({100.0*new_otif/max(len(new),1):5.1f}%)")

    old_i = old.drop_duplicates("_k").set_index("_k")
    new_i = new.drop_duplicates("_k").set_index("_k")
    common = old_i.index.intersection(new_i.index)
    only_old = old_i.index.difference(new_i.index)
    only_new = new_i.index.difference(old_i.index)

    print()
    print(f"lines in both     : {len(common):>7,}")
    print(f"only in workbook  : {len(only_old):>7,}")
    print(f"only in new run   : {len(only_new):>7,}   <- rows the old report was missing")
    if len(only_old):
        print(f"   sample only-in-workbook: {list(only_old[:5])}")
    if len(only_new):
        print(f"   sample only-in-new     : {list(only_new[:5])}")

    o, n = old_i.loc[common], new_i.loc[common]
    print()
    print("Field-level agreement on the lines present in both:")
    print(f"{'field':<32}{'agree':>9}{'differ':>9}   note")
    print("-" * 74)

    def report(field: str, agree: pd.Series, note: str = "") -> None:
        n_agree = int(agree.sum())
        n_diff = len(agree) - n_agree
        print(f"{field:<32}{n_agree:>9,}{n_diff:>9,}   {note}")

    report("OTIF", o["OTIF"].to_numpy() == n["OTIF"].to_numpy())
    report("MTA", o["MTA"].to_numpy() == n["MTA"].to_numpy())

    qty_close = (
        (o["Qty_Delivered_Before_Deadline"].fillna(0).to_numpy()
         - n["Qty_Delivered_Before_Deadline"].fillna(0).to_numpy()).round(3) == 0
    )
    report("Qty_Delivered_Before_Deadline", pd.Series(qty_close))

    both_na = o["Final_Deadline"].isna().to_numpy() & n["Final_Deadline"].isna().to_numpy()
    same = (o["Final_Deadline"].to_numpy() == n["Final_Deadline"].to_numpy()) | both_na
    report("Final_Deadline", pd.Series(same),
           "differences expected where RDD was blank")

    otif_diff = o.index[o["OTIF"].to_numpy() != n["OTIF"].to_numpy()]
    if len(otif_diff):
        print()
        print(f"Sample of the {len(otif_diff):,} OTIF disagreements:")
        cols = ["Quantity", "MTA", "Final_Deadline", "Qty_Delivered_Before_Deadline", "OTIF"]
        sample = otif_diff[:8]
        for key in sample:
            ov, nv = o.loc[key], n.loc[key]
            print(f"  {key}")
            print(f"     old: " + "  ".join(f"{c}={ov[c]!r}" for c in cols))
            print(f"     new: " + "  ".join(f"{c}={nv[c]!r}" for c in cols))

    print()
    print("Differences are expected in three places (see README): blank-date "
          "handling, rows the old key collision dropped, and rows beyond the old "
          "first-page-only read.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
