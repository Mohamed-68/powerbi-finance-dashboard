from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTHS)}


def looks_like_account_code(x: object) -> bool:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return False
    s = str(x).strip()
    return bool(re.fullmatch(r"\d{3,6}", s))  # 4001, 6002, 7001 etc.


def parse_messy_number(x: object) -> Optional[float]:
    """
    Handles:  "12,345" -> 12345
              "(2,119,020)" -> -2119020
              "-" or "" -> None
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


def find_first_data_row(raw: pd.DataFrame, account_col_idx: int = 0, max_scan: int = 80) -> Optional[int]:
    """
    Find first row where column 0 looks like an account code (4001, 5002...).
    """
    n = min(len(raw), max_scan)
    for i in range(n):
        if looks_like_account_code(raw.iat[i, account_col_idx]):
            return i
    return None


def infer_column_names(ncols: int) -> List[str]:
    """
    Your messy P&L is typically:
    0 Account | 1 Line Item | 2..13 Jan..Dec | 14 Total  (15 columns)
    """
    if ncols == 15:
        return ["Account", "Line Item"] + MONTHS + ["Total"]
    if ncols == 14:
        return ["Account", "Line Item"] + MONTHS  # no Total
    # fallback: still name the first 2 and keep the rest generic
    names = ["Account", "Line Item"]
    for i in range(2, ncols):
        names.append(f"Col{i+1}")
    return names


def clean_wide_pnl(raw: pd.DataFrame, year: int, scenario: str) -> Tuple[pd.DataFrame, Dict]:
    """
    Returns:
      - wide_clean: rows starting from first account row, with inferred headers
      - info: useful debug summary
    """
    start = find_first_data_row(raw, account_col_idx=0)
    info: Dict = {"first_data_row_index": start, "ncols": raw.shape[1], "nrows_raw": raw.shape[0]}

    if start is None:
        return pd.DataFrame(), {**info, "error": "Could not find an account-code row (e.g., 4001) in first 80 rows."}

    wide = raw.iloc[start:].copy()
    wide.columns = infer_column_names(raw.shape[1])

    # Keep only non-empty rows (some are fully blank)
    wide = wide.dropna(how="all")

    # Identify row types
    wide["IsAccountRow"] = wide["Account"].apply(looks_like_account_code)

    # Parse numeric values in month columns (and Total if present)
    num_cols = [c for c in MONTHS if c in wide.columns] + (["Total"] if "Total" in wide.columns else [])
    for c in num_cols:
        wide[c] = wide[c].apply(parse_messy_number)

    # Simple counts
    info.update({
        "rows_after_start": int(len(wide)),
        "account_rows": int(wide["IsAccountRow"].sum()),
        "non_account_rows": int((~wide["IsAccountRow"]).sum()),
        "detected_month_cols": [c for c in MONTHS if c in wide.columns],
        "has_total": "Total" in wide.columns,
        "sample_accounts": wide.loc[wide["IsAccountRow"], "Account"].astype(str).head(5).tolist(),
        "sample_lines": wide.loc[wide["IsAccountRow"], "Line Item"].astype(str).head(5).tolist(),
    })

    # Add scenario/year for context (even in wide)
    wide["Scenario"] = scenario
    wide["Year"] = year

    return wide, info


def wide_to_long_fact(wide: pd.DataFrame) -> pd.DataFrame:
    if wide.empty:
        return wide

    # Keep only true account rows (drop section headers + subtotals to avoid double counting)
    wide_acc = wide.loc[wide["IsAccountRow"]].copy()

    month_cols = [m for m in MONTHS if m in wide_acc.columns]
    long = wide_acc.melt(
        id_vars=["Account", "Line Item", "Scenario", "Year"],
        value_vars=month_cols,
        var_name="Month",
        value_name="Amount"
    )

    # Add date columns
    long["MonthNumber"] = long["Month"].map(MONTH_TO_NUM).astype("Int64")
    long["MonthStartDate"] = pd.to_datetime(
        long["Year"].astype(str) + "-" + long["MonthNumber"].astype(str) + "-01",
        errors="coerce"
    )
    long["MonthEndDate"] = (long["MonthStartDate"] + pd.offsets.MonthEnd(0)).dt.date
    long["Quarter"] = "Q" + (((long["MonthNumber"].astype(float) - 1) // 3) + 1).astype("Int64").astype(str)

    return long


def detect_pnl_sheets(xls_path: Path) -> List[Tuple[str, str, int]]:
    """
    Returns: list of (sheet_name, scenario, year)
    Matches names like: "Actual P&L 2024" or "Budget P&L 2021"
    """
    xls = pd.ExcelFile(xls_path)
    out = []
    pat = re.compile(r"^(Actual|Budget)\s+P&L\s+(\d{4})$", re.IGNORECASE)
    for s in xls.sheet_names:
        m = pat.match(s.strip())
        if m:
            scenario = m.group(1).title()
            year = int(m.group(2))
            out.append((s, scenario, year))
    return out


def main(xls_path: Path, out_dir: Path):
    out_dir.mkdir(parents=True, exist_ok=True)

    sheets = detect_pnl_sheets(xls_path)
    if not sheets:
        raise ValueError("No sheets matched pattern like 'Actual P&L 2024' / 'Budget P&L 2024'.")

    all_infos = []
    all_long = []

    for sheet_name, scenario, year in sheets:
        raw = pd.read_excel(xls_path, sheet_name=sheet_name, header=None, dtype=object)
        wide, info = clean_wide_pnl(raw, year=year, scenario=scenario)
        info["sheet_name"] = sheet_name
        all_infos.append(info)

        # Export wide for inspection
        wide_out = out_dir / f"wide_fixed_{sheet_name.replace('/','-')}.csv"
        wide.to_csv(wide_out, index=False)

        # Export “non-account rows” separately (subtotals/sections)
        if not wide.empty and "IsAccountRow" in wide.columns:
            non_acc = wide.loc[~wide["IsAccountRow"]].copy()
            non_acc_out = out_dir / f"wide_nonaccount_rows_{sheet_name.replace('/','-')}.csv"
            non_acc.to_csv(non_acc_out, index=False)

        # Long fact
        long = wide_to_long_fact(wide)
        all_long.append(long)

    # Combine
    fact = pd.concat(all_long, ignore_index=True)
    fact_out = out_dir / "FactPnL_Long_AllYears.csv"
    fact.to_csv(fact_out, index=False)

    # Summary
    summary = pd.DataFrame(all_infos).sort_values(["sheet_name"])
    summary_out = out_dir / "pnl_structure_summary.csv"
    summary.to_csv(summary_out, index=False)

    # Also write a simple markdown summary (easy to read)
    md = []
    md.append(f"# P&L Structure Summary\n")
    md.append(f"**File:** `{xls_path.as_posix()}`\n")
    md.append(f"**Generated:**\n- `{fact_out.as_posix()}`\n- `{summary_out.as_posix()}`\n")
    md.append("\n## Sheets\n")
    for _, row in summary.iterrows():
        md.append(f"### {row['sheet_name']}\n")
        md.append(f"- first_data_row_index: {row.get('first_data_row_index')}\n")
        md.append(f"- ncols: {row.get('ncols')}, nrows_raw: {row.get('nrows_raw')}\n")
        md.append(f"- account_rows: {row.get('account_rows')}, non_account_rows: {row.get('non_account_rows')}\n")
        md.append(f"- detected_month_cols: {row.get('detected_month_cols')}\n")
        md.append(f"- has_total: {row.get('has_total')}\n")
        md.append(f"- sample_accounts: {row.get('sample_accounts')}\n")
        md.append(f"- sample_lines: {row.get('sample_lines')}\n")

    (out_dir / "pnl_structure_summary.md").write_text("\n".join(md), encoding="utf-8")

    print("✅ Done.")
    print(f"- Fact table: {fact_out}")
    print(f"- Structure summary: {summary_out}")
    print(f"- Readable summary: {out_dir / 'pnl_structure_summary.md'}")
    print("\nTip: open the wide_fixed_*.csv files to see exactly how columns map (Account, Line Item, Jan..Dec, Total).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("excel_file", type=str, help="Path to your Excel workbook (e.g. data/Financial_Raw_Data_2020_2024.xlsx)")
    parser.add_argument("--out", type=str, default="reports", help="Output folder for reports/csvs")
    args = parser.parse_args()

    main(Path(args.excel_file), Path(args.out))
