from __future__ import annotations

import json
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# -----------------------------
# CONFIG
# -----------------------------
MONTH_NAMES = {"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"}
DEFAULT_FILE = Path("data/Financial_Raw_Data_2020_2024.xlsx")  # change if needed

REPORTS_DIR = Path("reports")
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

pd.set_option("display.max_columns", 50)
pd.set_option("display.width", 180)


# -----------------------------
# HELPERS
# -----------------------------
def normalize_text(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return str(x).strip()


def looks_like_account_code(x: str) -> bool:
    x = x.strip()
    return bool(re.fullmatch(r"\d{3,6}", x))  # e.g., 4001, 6002, 7001


def is_blank_col(col: pd.Series) -> bool:
    return col.isna().all() or (col.astype(str).str.strip() == "").all()


def row_score_for_header(row_values: List[str]) -> int:
    """
    Score a row as a potential header row.
    We expect something like: Account | Line Item | Jan | Feb | ... | Dec | Total
    """
    upper = [v.upper() for v in row_values]
    score = 0

    if any(v == "ACCOUNT" for v in upper):
        score += 4
    if any(v.replace(" ", "") in {"LINEITEM", "LINEITEMS"} or v == "LINE ITEM" for v in upper):
        score += 4

    # Count month-like cells
    months_found = sum(1 for v in upper if v in MONTH_NAMES)
    score += months_found * 1

    # Total column often exists
    if any(v == "TOTAL" for v in upper):
        score += 2

    return score


def detect_header_row(raw_df: pd.DataFrame, search_rows: int = 60) -> Tuple[Optional[int], Dict]:
    """
    Find the best header row index by scoring the first N rows.
    Returns (best_row_index, debug_info).
    """
    n = min(search_rows, len(raw_df))
    scores = []
    for i in range(n):
        row = raw_df.iloc[i].tolist()
        row_vals = [normalize_text(x) for x in row]
        score = row_score_for_header(row_vals)
        if score > 0:
            scores.append((i, score, row_vals[:15]))  # keep some preview
    scores_sorted = sorted(scores, key=lambda t: t[1], reverse=True)

    best = scores_sorted[0][0] if scores_sorted else None
    debug = {
        "candidates_top10": scores_sorted[:10],
        "best_row_index": best,
        "best_score": scores_sorted[0][1] if scores_sorted else None,
    }
    return best, debug


def extract_table_using_header(raw_df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """
    Use a detected header row to build a table.
    """
    df = raw_df.copy()
    df.columns = [normalize_text(x) for x in df.iloc[header_row].tolist()]
    df = df.iloc[header_row + 1 :].reset_index(drop=True)

    # Drop fully blank rows
    df = df.dropna(how="all")

    # Drop fully blank columns
    blank_cols = [c for c in df.columns if is_blank_col(df[c])]
    if blank_cols:
        df = df.drop(columns=blank_cols)

    return df


def find_month_columns(columns: List[str]) -> List[str]:
    months = []
    for c in columns:
        cu = str(c).strip().upper()
        if cu in MONTH_NAMES:
            months.append(c)
    return months


def profile_dataframe(df: pd.DataFrame, max_examples: int = 3) -> Dict:
    """
    Basic profiling: null%, unique, example values per column.
    """
    prof = {}
    for c in df.columns:
        s = df[c]
        null_pct = float(s.isna().mean())
        uniq = int(s.nunique(dropna=True))
        examples = (
            s.dropna()
             .astype(str)
             .map(lambda x: x.strip())
             .loc[lambda x: x != ""]
             .unique()[:max_examples]
             .tolist()
        )
        prof[c] = {"null_pct": null_pct, "unique": uniq, "examples": examples}
    return prof


def detect_pnl_row_types(df: pd.DataFrame, account_col: str, line_col: str) -> Dict[str, int]:
    """
    Count likely line types: account rows, section header rows, subtotal rows.
    """
    acc = df[account_col].astype(str).map(lambda x: x.strip())
    line = df[line_col].astype(str).map(lambda x: x.strip())

    is_account = acc.map(looks_like_account_code)
    is_section = (~is_account) & line.str.isupper() & (line != "")  # like "REVENUE"
    is_subtotal = (~is_account) & line.str.contains(r"^Total|Gross Profit|EBITDA|Net Income", case=False, regex=True)

    return {
        "rows_total": int(len(df)),
        "rows_account": int(is_account.sum()),
        "rows_section_headers": int(is_section.sum()),
        "rows_subtotals": int(is_subtotal.sum()),
        "rows_other_nonaccount": int((~is_account).sum()),
    }


# optional: numeric parser (useful later, but also helps you understand mess)
def parse_messy_number(x) -> Optional[float]:
    """
    Convert values like:
      "12,345" -> 12345
      "(2,119,020)" -> -2119020
      "-" or "" -> None
    Leaves numeric as float.
    """
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    if s in {"", "-"}:
        return None

    is_parens = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace(",", "").strip()

    try:
        val = float(s)
    except ValueError:
        return None

    return -val if is_parens else val


# -----------------------------
# MAIN AUDIT
# -----------------------------
@dataclass
class SheetAudit:
    sheet_name: str
    raw_shape: Tuple[int, int]
    detected_header_row: Optional[int]
    header_debug: Dict
    extracted_shape: Optional[Tuple[int, int]]
    columns: List[str]
    month_columns: List[str]
    blank_columns_removed: List[str]
    pnl_row_type_counts: Optional[Dict[str, int]]
    profile: Dict


def audit_sheet(xls_path: Path, sheet_name: str) -> SheetAudit:
    raw = pd.read_excel(xls_path, sheet_name=sheet_name, header=None, dtype=object)
    raw_shape = raw.shape

    header_row, debug = detect_header_row(raw, search_rows=60)

    extracted_shape = None
    cols: List[str] = []
    month_cols: List[str] = []
    removed_blank_cols: List[str] = []
    pnl_counts = None
    prof = {}

    if header_row is not None:
        # identify blank columns that would be removed
        temp = raw.copy()
        temp.columns = [f"Column{i+1}" for i in range(temp.shape[1])]
        removed_blank_cols = [c for c in temp.columns if is_blank_col(temp[c])]

        df = extract_table_using_header(raw, header_row=header_row)
        extracted_shape = df.shape
        cols = list(df.columns)
        month_cols = find_month_columns(cols)

        # profile
        prof = profile_dataframe(df)

        # if looks like a P&L (has Account + Line Item)
        account_col = next((c for c in cols if c.strip().lower() == "account"), None)
        line_col = next((c for c in cols if c.strip().lower().replace(" ", "") in {"lineitem","lineitems"} or c.strip().lower() == "line item"), None)
        if account_col and line_col:
            pnl_counts = detect_pnl_row_types(df, account_col=account_col, line_col=line_col)

    return SheetAudit(
        sheet_name=sheet_name,
        raw_shape=raw_shape,
        detected_header_row=header_row,
        header_debug=debug,
        extracted_shape=extracted_shape,
        columns=cols,
        month_columns=month_cols,
        blank_columns_removed=removed_blank_cols,
        pnl_row_type_counts=pnl_counts,
        profile=prof,
    )


def export_previews(xls_path: Path, sheet_names: List[str], rows: int = 25):
    """
    Export raw previews (first N rows with header=None) for selected sheets.
    This is what helps you understand why Power Query gets confused.
    """
    for s in sheet_names:
        raw = pd.read_excel(xls_path, sheet_name=s, header=None, dtype=object)
        out = REPORTS_DIR / f"preview_raw_{safe_name(s)}.csv"
        raw.head(rows).to_csv(out, index=False)
    print(f"Saved raw previews to: {REPORTS_DIR.resolve()}")


def safe_name(s: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()


def main(xls_path: Path):
    if not xls_path.exists():
        raise FileNotFoundError(f"Excel file not found: {xls_path.resolve()}")

    xls = pd.ExcelFile(xls_path)
    sheets = xls.sheet_names

    # 1) inventory
    inv = pd.DataFrame({"sheet_name": sheets})
    inv.to_csv(REPORTS_DIR / "sheet_inventory.csv", index=False)

    # 2) audit all sheets (or you can filter)
    audits: List[SheetAudit] = []
    for s in sheets:
        audits.append(audit_sheet(xls_path, s))

    # 3) export json report
    audit_json = [asdict(a) for a in audits]
    (REPORTS_DIR / "audit_report.json").write_text(json.dumps(audit_json, indent=2), encoding="utf-8")

    # 4) write a human-readable markdown summary
    md_lines = []
    md_lines.append(f"# Data Audit Report\n")
    md_lines.append(f"**File:** `{xls_path.as_posix()}`\n")
    md_lines.append(f"## Sheets\n")
    for a in audits:
        md_lines.append(f"### {a.sheet_name}\n")
        md_lines.append(f"- Raw shape: {a.raw_shape}\n")
        md_lines.append(f"- Detected header row (0-based): {a.detected_header_row}\n")
        if a.detected_header_row is not None:
            md_lines.append(f"- Extracted shape: {a.extracted_shape}\n")
            md_lines.append(f"- Columns: {', '.join(a.columns[:20])}{' ...' if len(a.columns) > 20 else ''}\n")
            md_lines.append(f"- Month columns: {', '.join(a.month_columns) if a.month_columns else 'None detected'}\n")
            if a.pnl_row_type_counts:
                md_lines.append(f"- P&L row types: {a.pnl_row_type_counts}\n")
        else:
            md_lines.append(f"- Header not detected (check raw preview)\n")

    (REPORTS_DIR / "audit_report.md").write_text("\n".join(md_lines), encoding="utf-8")

    # 5) export raw previews for the important sheets (helps debugging Power Query)
    key_sheets = [s for s in sheets if re.search(r"(Actual P&L 2024|Budget P&L 2024)", s, re.IGNORECASE)]
    if key_sheets:
        export_previews(xls_path, key_sheets, rows=35)

    print("✅ Audit complete.")
    print(f"- Inventory: {REPORTS_DIR / 'sheet_inventory.csv'}")
    print(f"- JSON: {REPORTS_DIR / 'audit_report.json'}")
    print(f"- Markdown: {REPORTS_DIR / 'audit_report.md'}")
    if key_sheets:
        print(f"- Raw previews: {REPORTS_DIR / 'preview_raw_*.csv'}")


if __name__ == "__main__":
    # Change this path if your file is elsewhere
    main(DEFAULT_FILE)
