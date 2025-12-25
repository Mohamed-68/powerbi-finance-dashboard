from pathlib import Path
import pandas as pd
import numpy as np
import re

FILE_PATH = Path("Financial_Raw_Data_2020_2024.xlsx")

# Skip the “weird” one you mentioned (support both naming styles)
SKIP_SHEETS = {
    "Actual_P_L_2021",
    "Actual P&L 2021",
}

# Heuristics: which sheets are likely “wide P&L statements”
PNL_NAME_HINTS = ("actual", "budget", "p&l", "p_l", "p l", "pnl")


def is_pnl_statement_sheet(sheet_name: str) -> bool:
    n = sheet_name.strip().lower()
    return ("actual" in n or "budget" in n) and any(h in n for h in PNL_NAME_HINTS)


def first_data_row_index(df_raw: pd.DataFrame) -> int | None:
    """
    Find the first row that looks like data (account code in col 0).
    Works even if the top rows are titles/notes.
    """
    if df_raw.shape[1] == 0:
        return None

    col0 = df_raw.iloc[:, 0]
    # Convert to numeric where possible
    num = pd.to_numeric(col0, errors="coerce")
    # First row where account becomes a number (e.g., 4001, 5001, etc.)
    idx = num.first_valid_index()
    return int(idx) if idx is not None else None


def preview_top_lines(df_raw: pd.DataFrame, max_lines: int = 3) -> list[str]:
    """
    For messy P&L sheets, top lines are usually title/currency/note lines stored in first column.
    """
    lines = []
    for i in range(min(max_lines, len(df_raw))):
        val = df_raw.iloc[i, 0]
        if pd.isna(val):
            continue
        text = str(val).strip()
        if text:
            lines.append(text)
    return lines


def detect_header_row_structured(df_raw: pd.DataFrame, scan_rows: int = 30) -> int:
    """
    Try to detect which row is the header for structured tables.
    Default to row 0 if nothing obvious is found.
    """
    expected_tokens = {
        "account", "scenario", "amount", "month", "monthend", "monthenddate",
        "date", "year", "quarter", "department", "product", "entity"
    }
    scan_rows = min(scan_rows, len(df_raw))

    for r in range(scan_rows):
        row = df_raw.iloc[r].astype(str).str.lower().fillna("")
        hits = 0
        for t in expected_tokens:
            if row.str.contains(re.escape(t)).any():
                hits += 1
        # If we hit at least 2 expected tokens, treat it as a header row
        if hits >= 2:
            return r

    return 0


def main() -> None:
    if not FILE_PATH.exists():
        raise FileNotFoundError(f"Excel file not found: {FILE_PATH.resolve()}")

    xl = pd.ExcelFile(FILE_PATH)
    sheets = xl.sheet_names

    report_lines = []
    report_lines.append(f"# Sheet Header Report for `{FILE_PATH.name}`\n")
    report_lines.append(f"Total sheets: **{len(sheets)}**\n")

    for s in sheets:
        if s in SKIP_SHEETS:
            report_lines.append(f"## {s}\n")
            report_lines.append("**SKIPPED** (as requested)\n")
            continue

        # Read a raw chunk to inspect layout safely
        df_raw = pd.read_excel(FILE_PATH, sheet_name=s, header=None, nrows=60)

        report_lines.append(f"## {s}\n")

        if is_pnl_statement_sheet(s):
            report_lines.append("Type: **Wide P&L Statement (raw / messy)**\n")

            top_lines = preview_top_lines(df_raw, max_lines=3)
            if top_lines:
                report_lines.append("Top header lines (column A):\n")
                for t in top_lines:
                    report_lines.append(f"- {t}\n")
                report_lines.append("\n")

            start_idx = first_data_row_index(df_raw)
            if start_idx is None:
                report_lines.append("Could not detect first data row (no numeric Account found in col A).\n\n")
            else:
                report_lines.append(f"Detected first data row (0-based index): **{start_idx}**  \n")
                report_lines.append(f"Detected first data row (Excel row number): **{start_idx + 1}**\n\n")

                # Show a small preview from where the data begins
                preview = df_raw.iloc[start_idx:start_idx + 5, :8].copy()
                preview.columns = [f"Col{c+1}" for c in range(preview.shape[1])]
                report_lines.append("Preview (first 5 rows, first 8 cols from detected data start):\n\n")
                report_lines.append(preview.to_markdown(index=False))
                report_lines.append("\n\n")

        else:
            report_lines.append("Type: **Structured table**\n")
            header_row = detect_header_row_structured(df_raw)

            # Re-read using detected header row
            df = pd.read_excel(FILE_PATH, sheet_name=s, header=header_row, nrows=5)

            report_lines.append(f"Detected header row (0-based): **{header_row}**  \n")
            report_lines.append(f"Detected header row (Excel row): **{header_row + 1}**\n\n")
            report_lines.append("Columns:\n")
            for c in df.columns:
                report_lines.append(f"- {c}\n")
            report_lines.append("\n")
            report_lines.append("Preview (first 5 rows):\n\n")
            report_lines.append(df.to_markdown(index=False))
            report_lines.append("\n\n")

    out_path = Path("sheet_headers_report.md")
    out_path.write_text("".join(report_lines), encoding="utf-8")

    print(f"✅ Report written to: {out_path.resolve()}")
    print("Open `sheet_headers_report.md` to review all sheet header info.")


if __name__ == "__main__":
    main()
