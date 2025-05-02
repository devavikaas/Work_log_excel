import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from datetime import datetime

# === Configuration ===
start_date = datetime(2025, 5, 2)
end_date = datetime(2025, 5, 25)
date_range = pd.date_range(start=start_date, end=end_date)

columns = [
    "Start Time", "End Time", "Duration (minutes)",
    "Type", "Task Name", "Description / Subject", "Remarks"
]

header_colors = [
    "FBE5D6", "D9EAD3", "D9D2E9", "CFE2F3", "FFF2CC", "F4CCCC", "EAD1DC"
]
header_fills = [
    PatternFill(start_color=color, end_color=color, fill_type="solid")
    for color in header_colors
]
header_font = Font(bold=True, size=16)
cell_font = Font(size=14)
wrap_alignment = Alignment(wrap_text=True, vertical="top")

summary_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
tab_colors = [
    "1072BA", "A9D08E", "FFD966", "D9D2E9", "F4CCCC", "B4C7E7", "FFE699",
    "C9DAF8", "D5A6BD", "B6D7A8", "EAD1DC", "F9CB9C", "B7E1CD", "F6B26B",
    "D0E0E3", "FCE5CD", "A2C4C9", "C27BA0", "A4C2F4", "D5A6BD", "B6D7A8",
    "D9D2E9", "C9DAF8", "F4CCCC"
]

file_path = "/content/Work_Study_Log_May_2025_Final.xlsx"

# === Step 1: Create base workbook ===
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    for date in date_range:
        sheet_name = date.strftime("%Y-%m-%d")
        df = pd.DataFrame([[""] * len(columns)], columns=columns)
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# === Step 2: Style and formula per sheet ===
wb = openpyxl.load_workbook(file_path)

for i, sheet_name in enumerate(wb.sheetnames):
    if sheet_name == "Summary":
        continue

    ws = wb[sheet_name]

    # Styled headers
    for col_idx, (fill, header) in enumerate(zip(header_fills, columns), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Apply font size 14 and word wrap to all data cells
    for row in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=len(columns)):
        for cell in row:
            cell.font = cell_font
            cell.alignment = wrap_alignment

    # Manual formulas for Start Time and Duration
    for row in range(2, 100):
        start_cell = f"A{row+1}"
        end_cell = f"B{row}"
        duration_cell = f"C{row}"
        ws[start_cell] = f"={end_cell}"
        ws[duration_cell] = (
            f"=IF(AND(ISNUMBER(B{row}), ISNUMBER(A{row})), ROUND((B{row}-A{row})*1440, 0), \"\")"
        )

    # Format time columns
    for col in ["A", "B"]:
        for row in range(2, 100):
            ws[f"{col}{row}"].number_format = "h:mm:ss AM/PM"

    # Duration as number
    for row in range(2, 100):
        ws[f"C{row}"].number_format = "0"

    # Dropdown in "Type" column
    dv = DataValidation(type="list", formula1='"Task,Break"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("D2:D100")

    # Add table
    end_col_letter = get_column_letter(len(columns))
    table_range = f"A1:{end_col_letter}100"
    table_name = "T_" + sheet_name.replace("-", "_")
    table = Table(displayName=table_name, ref=table_range)
    table_style = TableStyleInfo(name="TableStyleMedium6", showRowStripes=True)
    table.tableStyleInfo = table_style
    ws.add_table(table)

    # Summary block on right
    ws["I1"] = "Total Study Time"
    ws["J1"] = "=SUMIF(D2:D100, \"Task\", C2:C100)/1440"
    ws["I2"] = "Total Break Time"
    ws["J2"] = "=SUMIF(D2:D100, \"Break\", C2:C100)/1440"
    ws["I3"] = "Focus Ratio"
    ws["J3"] = "=IF((J1+J2)=0, 0, J1/(J1+J2))"

    ws["J1"].number_format = "hh:mm:ss"
    ws["J2"].number_format = "hh:mm:ss"
    ws["J3"].number_format = "0.00%"

    for cell in ["I1", "I2", "I3"]:
        ws[cell].font = header_font
        ws[cell].fill = summary_fill
        ws[cell].alignment = Alignment(horizontal="left", wrap_text=True)

    for cell in ["J1", "J2", "J3"]:
        ws[cell].font = cell_font
        ws[cell].alignment = wrap_alignment

    # Auto-adjust column widths
    for col_idx in range(1, len(columns) + 1):
        col_letter = get_column_letter(col_idx)
        max_length = max(
            len(str(ws[f"{col_letter}{row}"].value) or "")
            for row in range(1, 101)
        )
        ws.column_dimensions[col_letter].width = max_length + 2

    # Tab color
    ws.sheet_properties.tabColor = tab_colors[i % len(tab_colors)]

# === Step 3: Create Summary Sheet ===
summary = wb.create_sheet("Summary")
summary_headers = ["Day", "Total Study Time", "Total Break Time", "Focus Ratio"]

for col_idx, header in enumerate(summary_headers, 1):
    cell = summary.cell(row=1, column=col_idx, value=header)
    cell.font = Font(bold=True, size=16)
    cell.fill = summary_fill
    cell.alignment = Alignment(horizontal="center", wrap_text=True)

row = 2
for sheet in wb.sheetnames:
    if sheet == "Summary":
        continue
    sheet_ref = f"'{sheet}'"
    summary[f"A{row}"] = sheet
    summary[f"B{row}"] = f"=SUMIF({sheet_ref}!D2:D100, \"Task\", {sheet_ref}!C2:C100)/1440"
    summary[f"C{row}"] = f"=SUMIF({sheet_ref}!D2:D100, \"Break\", {sheet_ref}!C2:C100)/1440"
    summary[f"D{row}"] = f"=IF((B{row}+C{row})=0, 0, B{row}/(B{row}+C{row}))"

    summary[f"B{row}"].number_format = "hh:mm:ss"
    summary[f"C{row}"].number_format = "hh:mm:ss"
    summary[f"D{row}"].number_format = "0.00%"
    row += 1

# Font size 14 and word wrap for summary data cells
for row_cells in summary.iter_rows(min_row=2, max_row=row-1, min_col=1, max_col=4):
    for cell in row_cells:
        cell.font = cell_font
        cell.alignment = wrap_alignment

# Auto-adjust column widths in summary
for col_idx in range(1, 5):
    col_letter = get_column_letter(col_idx)
    max_length = max(
        len(str(summary[f"{col_letter}{row}"].value) or "")
        for row in range(1, row)
    )
    summary.column_dimensions[col_letter].width = max_length + 2

# Apply table styling to summary
summary_range = f"A1:D{row-1}"
summary_table = Table(displayName="SummaryTable", ref=summary_range)
summary_style = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
summary_table.tableStyleInfo = summary_style
summary.add_table(summary_table)

# Save workbook
wb.save(file_path)

# Optional: For Colab download
from google.colab import files
files.download(file_path)
