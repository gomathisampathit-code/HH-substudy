import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl.styles import Alignment, Border, Side, Font
from io import BytesIO
from openpyxl import Workbook

# --- PASSWORD PROTECTION ---
st.title("🔒 RespIndNet HH Substudy Specimen Transfer Form (Virology)")
password = st.text_input("Enter Password:", type="password")

if password != "HH123": 
    st.warning("Please enter the correct password.")
    st.stop()

# --- Load Google Sheet ---
# HH Sample Collection Form
sheet_id = "1wZNK_uRuTlFWtS4HAS-5dvxfMZzhe0u4vwgBWkvKD-4"
csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv"
df = pd.read_csv(csv_url, on_bad_lines="skip")


# Normalize
df.columns = df.columns.str.strip().str.lower()
df["submissiondate"] = pd.to_datetime(df["submissiondate"], errors="coerce")
df["submissiondate"] = df["submissiondate"].dt.tz_localize(None)
df["dt_sample"] = pd.to_datetime(df["dt_sample"], errors="coerce")
df["dt_sample"] = df["dt_sample"].dt.tz_localize(None)

# Clean episode1
df["episode1"] = df["episode1"].fillna("").astype(str).str.strip()

# --- Filter to today's collected, virology-ready samples ---
# sample_collected == 1 -> collected today AND always a Respiratory swab for this sub-study
df["sample_collected"] = df["sample_collected"].astype(str).str.strip().str.lower()
today_str = pd.Timestamp.today().strftime("%Y-%m-%d")
df_today = df[
    (df["submissiondate"].dt.strftime("%Y-%m-%d") == today_str)
    & (df["sample_collected"] == "yes")
].copy()
df_today = df_today.reset_index(drop=True)

# Get existing episode1 for each child_id
episode_lookup = (
    df[df["episode1"] != ""]
    .drop_duplicates("child_id")
    .set_index("child_id")["episode1"]
)

# Display only existing episode IDs
df_today["episode_display"] = (
    df_today["child_id"]
    .map(episode_lookup)
    .fillna("")
)

# Sample type is always Respiratory swab when sample_collected == 1
df_today["type_of_sample"] = "Respiratory swab"

# Barcode ID
df_today["barcode_id"] = df_today["sample_id"]

# Sample sequence per household member (S.PER IND)
df_today["sample_sequence"] = df_today.groupby("member_id").cumcount() + 1

def split_member(val):
    if pd.isna(val) or val == "":
        return "", ""
    parts = str(val).split(",", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return parts[0].strip(), ""

df_today["member_id_clean"], df_today["member_name"] = zip(
    *df_today["member_id"].map(split_member)
)

df_today["index_case_label"] = (
    pd.to_numeric(df_today["index_case"], errors="coerce")
    .map({
        1: "Index",
        2: "Contact"
    })
    .fillna("")
)

# Final table
table = pd.DataFrame({
    "S.NO": range(1, len(df_today) + 1),
    "BARCODE ID": df_today["barcode_id"],
    "Episode": df_today["episode_display"],
    "Index case": df_today["index_case_label"],
    "S.TYPE": df_today["type_of_sample"],
    "S.PER IND": df_today["sample_sequence"],
    "HH substudy member ID": df_today["member_id_clean"],
    "NAME": df_today["member_name"],
    "Day": df_today.get("sample_timepoint", ""),
    "S.C DATE/TIME": df_today["dt_sample"],
    "STUDY": "",
    "RECEIVED BY": "",
    "VOL (VIRO)": "",
    "REMARKS(LYSED/LIPEMIC/ICTERIC/SAMPLE SPILLAGE)": ""
})

st.subheader("📋 Generated Table")
st.dataframe(table)

# --- Download Excel ---
if len(table) > 0:
    today_str_display = datetime.today().strftime("%d-%m-%Y")
    excel_filename = f"{today_str_display}_RespIndNet_HH_STF(Field_to_Virology).xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Header row 1 (merged)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=table.shape[1])
    cell = ws.cell(row=1, column=1, value="RespIndNet_HH substudy Specimen Transfer Form (Virology)")
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")

    # Apply border to all cells in the merged header row
    for col in range(1, table.shape[1] + 1):
        ws.cell(row=1, column=col).border = border

    # Extra span row (row 2)
    spans = [
        "Sample Shipment Date and Time:",
        "Field Manager Sign/Initials:",
        "Virology Staff Sign/Initials:",
        "To be filled by Virology"
    ]

    col_split = [
        table.shape[1] // 4,
        table.shape[1] // 4,
        table.shape[1] - 3 * (table.shape[1] // 4),
        table.shape[1] // 4
    ]

    start_col = 1
    for i, val in enumerate(spans):
        end_col = start_col + col_split[i] - 1
        ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
        cell = ws.cell(row=2, column=start_col, value=val)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for row in range(2, 3):
            for col in range(start_col, end_col + 1):
                ws.cell(row=row, column=col).border = border
        start_col = end_col + 1

    # Column headers (row 3)
    for j, col_name in enumerate(table.columns, 1):
        c = ws.cell(row=3, column=j, value=col_name)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    # Data rows
    for i, row in table.iterrows():
        for j, val in enumerate(row, 1):
            c = ws.cell(row=i + 4, column=j, value=val)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = border

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download Excel",
        data=buffer,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
