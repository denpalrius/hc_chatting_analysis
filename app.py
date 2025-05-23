
from io import BytesIO
from datetime import datetime
import streamlit as st
import pandas as pd
import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def get_individual_name(content: bytes, filename: str) -> str:
    """Extract the person’s name from D3 (xls or xlsx)."""
    if filename.lower().endswith(".xls"):
        wb = xlrd.open_workbook(file_contents=content)
        sh = wb.sheet_by_index(0)
        return sh.cell_value(2, 3).split(",")[0].strip()
    else:
        wb = openpyxl.load_workbook(BytesIO(content), read_only=True, data_only=True)
        raw = wb.worksheets[0].cell(row=3, column=4).value
        return str(raw).split(",")[0].strip()


def detect_header_row(raw: pd.DataFrame) -> int:
    """
    Locate the zero-based row index of the "Date" header that sits
    below "Time Zone:" and any ISP…LPN section header, regardless of
    what's between ISP and LPN.
    """
    col0 = raw.iloc[:, 0].astype(str)

    # find "Time Zone:"
    tz_mask = col0.str.contains("Time Zone:", case=False, na=False)
    tz_idx = tz_mask.idxmax() if tz_mask.any() else -1

    # match any header that starts with ISP and later contains LPN
    isp_mask = col0.str.contains(r"^ISP.*LPN", case=False, na=False)
    isp_idx = isp_mask.idxmax() if isp_mask.any() else -1

    # find all rows where column A literally equals "Date"
    date_mask = col0.str.strip().str.lower() == "date"
    candidates = raw.index[date_mask]

    # pick the first "Date" row below both markers
    threshold = max(tz_idx, isp_idx)
    for r in candidates:
        if r > threshold:
            return r

    raise RuntimeError("Could not locate the 'Date' header under the required section.")


def find_pdn_lpn_block_end(raw: pd.DataFrame, header_row: int) -> int:
    """
    Find the end of the PDN-LPN block by looking for the first "Total"
    in column C below the header row.
    """
    col2 = raw.iloc[:, 2].astype(str).str.strip().str.lower()
    total_mask = (col2 == "total") & (raw.index > header_row)
    end_candidates = raw.index[total_mask]
    return end_candidates[0] if len(end_candidates) > 0 else len(raw)


def parse_file(content: bytes, filename: str) -> pd.DataFrame:
    """
    Read one per-person sheet, isolate only the PDN-LPN block,
    then return rows with a real Date in col A plus Duration (minutes) & Provider.
    """
    indiv = get_individual_name(content, filename)
    raw = pd.read_excel(BytesIO(content), header=None)

    header_row = detect_header_row(raw)
    end_row = find_pdn_lpn_block_end(raw, header_row)

    section = raw.iloc[header_row + 1 : end_row]
    parsed = pd.to_datetime(section.iloc[:, 0], errors="coerce")
    valid = parsed.notna()

    df = section.loc[valid, [0, 4, 6]].copy()
    df.columns = ["Date", "Duration_min", "Service Provider"]
    df["Date"] = parsed[valid].dt.date
    df["Duration"] = df["Duration_min"].astype(float) / 60.0
    df["Individual"] = indiv

    return df[["Date", "Service Provider", "Individual", "Duration"]]


def format_duration(hours: float):
    """If whole hours, return int; otherwise 'H:MM'."""
    h = int(hours)
    m = int(round((hours - h) * 60))
    return h if m == 0 else f"{h}:{m:02d}"


def build_summary_workbook(df_raw: pd.DataFrame) -> Workbook:
    """
    Build the DailyMatrix sheet, using Excel formulas for all totals.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "DailyMatrix"
    bold = Font(bold=True)

    individuals = sorted(df_raw["Individual"].unique())
    total_col = len(individuals) + 2  # last column index for Provider Total

    row = 1
    for date in sorted(df_raw["Date"].unique()):
        day = df_raw[df_raw["Date"] == date]

        # Date header
        ws.cell(row=row, column=1, value=date.strftime("%m/%d/%Y")).font = bold
        row += 1

        # Column headers
        ws.cell(row=row, column=1, value="Service Provider").font = bold
        for idx, name in enumerate(individuals, start=2):
            ws.cell(row=row, column=idx, value=name).font = bold
        ws.cell(row=row, column=total_col, value="Provider Total").font = bold
        row += 1

        # mark where provider rows start
        provider_start = row

        # Provider rows
        for prov in day["Service Provider"].unique():
            ws.cell(row=row, column=1, value=prov)

            # write each individual's hours as plain number/text
            for idx, name in enumerate(individuals, start=2):
                hrs = day[
                    (day["Service Provider"] == prov) & (day["Individual"] == name)
                ]["Duration"].sum()
                ws.cell(row=row, column=idx, value=format_duration(hrs))

            # write an Excel SUM formula for this provider’s total across all individuals
            first_data_col = get_column_letter(2)
            last_data_col = get_column_letter(1 + len(individuals))
            ws.cell(
                row=row,
                column=total_col,
                value=f"=SUM({first_data_col}{row}:{last_data_col}{row})",
            ).font = bold

            row += 1

        # mark where provider rows end
        provider_end = row - 1

        # Totals per individual (row of sums)
        ws.cell(row=row, column=1, value="Total hours for individual").font = bold
        for idx, name in enumerate(individuals, start=2):
            col_letter = get_column_letter(idx)
            # sum all the cells in this column from provider_start to provider_end
            ws.cell(
                row=row,
                column=idx,
                value=f"=SUM({col_letter}{provider_start}:{col_letter}{provider_end})",
            ).font = bold
        row += 1

        # 24-hour cap remaining
        ws.cell(row=row, column=1, value="Total hrs pending in a 24hr period").font = (
            bold
        )
        for idx in range(2, 2 + len(individuals)):
            col_letter = get_column_letter(idx)
            # subtract the total-hours cell (just above) from 24
            ws.cell(row=row, column=idx, value=f"=24 - {col_letter}{row-1}").font = bold
        row += 2  # blank line before next date block

    return wb


def main():
    st.title("Daily Staffing Analysis App")

    files = st.file_uploader(
        "Upload your per-person Excel files (.xls/.xlsx).",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )
    if not files:
        return

    dfs = [parse_file(f.read(), f.name) for f in files]
    df_raw = pd.concat(dfs, ignore_index=True)

    wb = build_summary_workbook(df_raw)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("✅ PDN-LPN summary ready!")
    today = datetime.now().strftime("%Y-%m-%d")
    st.download_button(
        "Download Summary Excel",
        data=buf,
        file_name=f"daily_summary_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
