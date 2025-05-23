import streamlit as st
import pandas as pd
import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime

st.title("Daily Staffing Analysis App")


# File uploader
uploaded_files = st.file_uploader(
    "Upload the per-person Excel files for the billing period.",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    records = []
    for file in uploaded_files:
        content = file.read()

        # Extract individual name from D3
        if file.name.lower().endswith(".xls"):
            book = xlrd.open_workbook(file_contents=content)
            sh = book.sheet_by_index(0)
            file_ind = sh.cell_value(2, 3).split(",")[0].strip()
        else:
            wb2 = openpyxl.load_workbook(
                BytesIO(content), read_only=True, data_only=True
            )
            sh2 = wb2.worksheets[0]
            file_ind = str(sh2.cell(row=3, column=4).value).split(",")[0].strip()

        # Read rows from Excel row 41 onward
        df_temp = pd.read_excel(
            BytesIO(content),
            header=None,
            skiprows=40,
            usecols=[0, 3, 6],
            names=["Date", "Duration", "Service Provider"],
        )

        # Clean & parse
        df_temp = df_temp.dropna(subset=["Date", "Duration"])
        df_temp = df_temp[df_temp["Date"].astype(str) != "Date"]
        df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors="coerce").dt.date
        df_temp = df_temp.dropna(subset=["Date"])
        df_temp["Individual"] = file_ind
        df_temp["Duration"] = (
            df_temp["Duration"]
            .astype(str)
            .apply(lambda x: f"{x}:00" if len(x.split(":")) == 2 else x)
        )
        df_temp["Duration"] = (
            pd.to_timedelta(df_temp["Duration"]).dt.total_seconds() / 3600
        )

        records.append(df_temp)

    # Combine all into one DataFrame
    df_raw = pd.concat(records, ignore_index=True)

    # Build the summary workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "DailyMatrix"
    bold = Font(bold=True)
    row = 1
    individuals = sorted(df_raw["Individual"].unique())

    for current_date in sorted(df_raw["Date"].unique()):
        day_df = df_raw[df_raw["Date"] == current_date]

        # Date header
        ws.cell(row=row, column=1, value=current_date.strftime("%m/%d/%Y")).font = bold
        row += 1

        # Column headers
        ws.cell(row=row, column=1, value="Service Provider").font = bold
        for idx, indiv in enumerate(individuals, start=2):
            ws.cell(row=row, column=idx, value=indiv).font = bold
        total_col = 2 + len(individuals)
        ws.cell(row=row, column=total_col, value="Provider Total").font = bold
        row += 1

        provider_start = row

        # One row per provider
        for provider in day_df["Service Provider"].unique():
            ws.cell(row=row, column=1, value=provider)
            for idx, indiv in enumerate(individuals, start=2):
                hrs = day_df[
                    (day_df["Service Provider"] == provider)
                    & (day_df["Individual"] == indiv)
                ]["Duration"].sum()
                ws.cell(row=row, column=idx, value=hrs)

            # SUM formula for provider total
            c1 = get_column_letter(2)
            c2 = get_column_letter(1 + len(individuals))
            formula = f"=SUM({c1}{row}:{c2}{row})"
            cell = ws.cell(row=row, column=total_col, value=formula)
            cell.font = bold
            row += 1

        provider_end = row - 1

        # Totals per individual
        ws.cell(row=row, column=1, value="Total hours for individual").font = bold
        for idx in range(2, 2 + len(individuals)):
            col = get_column_letter(idx)
            formula = f"=SUM({col}{provider_start}:{col}{provider_end})"
            cell = ws.cell(row=row, column=idx, value=formula)
            cell.font = bold
        row += 1

        # Remaining hours to 24h cap
        ws.cell(row=row, column=1, value="Total hrs pending in a 24hr period").font = (
            bold
        )
        for idx in range(2, 2 + len(individuals)):
            col = get_column_letter(idx)
            above = f"{col}{row-1}"
            formula = f"=24 - {above}"
            cell = ws.cell(row=row, column=idx, value=formula)
            cell.font = bold
        row += 2

    # Prepare download
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.success("Summary workbook generated successfully!")

    today_date = datetime.now().strftime("%Y-%m-%d")

    st.download_button(
        label="Download Summary Excel",
        data=buf,
        file_name=f"daily_summary_{today_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
