# Work Log Excel

This Python script helps you track your work or study time in Excel. It automates the logging of start and end times, calculates durations, and generates daily summaries for your tasks and breaks.

## Features:
- Track study/work time and breaks.
- Automatically calculate time durations.
- Generate Excel reports with work session summaries.

README.md: Detailed Explanation of the Work Log Excel Code
markdown
Copy
Edit
# Work Log Excel – Time Tracking with Google Colab

This Python script generates an Excel workbook for tracking work and study sessions over a specific date range. It uses `pandas`, `openpyxl`, and Excel formulas to automate time logging, calculations, and formatting. The workbook includes detailed daily logs, summary reports, and automatic calculations for duration and focus ratio.

## Code Breakdown

### Import Libraries

```python
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from datetime import datetime
pandas: Used for date range generation and managing data.

openpyxl: Library for handling Excel file creation, modification, and formatting.

Font, PatternFill, Alignment: Used for customizing the fonts, fills, and alignments in the Excel sheet.

Table, TableStyleInfo: For creating tables with styles in the workbook.

DataValidation: For adding dropdown lists to specific cells (e.g., selecting between "Task" or "Break").

get_column_letter: Used to convert column indexes to Excel column letters.

datetime: For handling date and time operations.

Step 1: Configuration and Base Workbook Creation
python
Copy
Edit
start_date = datetime(2025, 5, 2)
end_date = datetime(2025, 5, 25)
date_range = pd.date_range(start=start_date, end=end_date)
start_date and end_date specify the range of dates for which the work log will be generated.

date_range generates a series of dates between start_date and end_date using pandas.date_range.

python
Copy
Edit
columns = [
    "Start Time", "End Time", "Duration (minutes)",
    "Type", "Task Name", "Description / Subject", "Remarks"
]
columns: Defines the column headers that will appear in the workbook for each work log entry.

python
Copy
Edit
header_colors = [
    "FBE5D6", "D9EAD3", "D9D2E9", "CFE2F3", "FFF2CC", "F4CCCC", "EAD1DC"
]
header_fills = [
    PatternFill(start_color=color, end_color=color, fill_type="solid")
    for color in header_colors
]
header_colors: List of colors to be used in the header row for each column.

header_fills: Applies the colors as PatternFill styles to the header cells.

python
Copy
Edit
header_font = Font(bold=True, size=16)
cell_font = Font(size=14)
wrap_alignment = Alignment(wrap_text=True, vertical="top")
header_font: Specifies the font style for headers (bold and size 16).

cell_font: Specifies the font style for data cells (size 14).

wrap_alignment: Defines alignment with word wrap and vertical alignment at the top of each cell.

python
Copy
Edit
summary_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
summary_fill: Defines a fill color for summary cells (light grey).

python
Copy
Edit
file_path = "/content/Work_Study_Log_May_2025_Final.xlsx"
file_path: Specifies the location where the final Excel file will be saved.

Step 2: Creating the Excel Workbook and Sheet Formatting
python
Copy
Edit
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    for date in date_range:
        sheet_name = date.strftime("%Y-%m-%d")
        df = pd.DataFrame([[""] * len(columns)], columns=columns)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
This block creates a new Excel workbook using pandas.ExcelWriter.

For each date in date_range, a new sheet is created with the date as the sheet name.

An empty DataFrame (df) is written to each sheet, with the defined columns.

Step 3: Applying Styles, Formulas, and Validation
python
Copy
Edit
wb = openpyxl.load_workbook(file_path)
Loads the Excel workbook created in the previous step using openpyxl.

python
Copy
Edit
for i, sheet_name in enumerate(wb.sheetnames):
    if sheet_name == "Summary":
        continue
    ws = wb[sheet_name]
Loops over all sheet names (except "Summary").

For each sheet, it accesses the corresponding ws (worksheet) object.

python
Copy
Edit
# Styled headers
for col_idx, (fill, header) in enumerate(zip(header_fills, columns), 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
For each column, the header cell is styled with the font, fill color, and alignment defined earlier.

python
Copy
Edit
# Apply font size 14 and word wrap to all data cells
for row in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=len(columns)):
    for cell in row:
        cell.font = cell_font
        cell.alignment = wrap_alignment
This applies the font style and word wrapping to all data cells in each row (starting from row 2).

Step 4: Adding Formulas and Data Validation
python
Copy
Edit
# Manual formulas for Start Time and Duration
for row in range(2, 100):
    start_cell = f"A{row+1}"
    end_cell = f"B{row}"
    duration_cell = f"C{row}"
    ws[start_cell] = f"={end_cell}"
    ws[duration_cell] = f"=IF(AND(ISNUMBER(B{row}), ISNUMBER(A{row})), ROUND((B{row}-A{row})*1440, 0), \"\")"
Adds formulas to calculate the Start Time and Duration (in minutes). The formula calculates the difference between start and end times, multiplying by 1440 to convert hours to minutes.

python
Copy
Edit
# Dropdown in "Type" column
dv = DataValidation(type="list", formula1='"Task,Break"', allow_blank=True)
ws.add_data_validation(dv)
dv.add("D2:D100")
Adds a dropdown list to the "Type" column (cells D2 to D100) for users to select "Task" or "Break".

python
Copy
Edit
# Add table
end_col_letter = get_column_letter(len(columns))
table_range = f"A1:{end_col_letter}100"
table_name = "T_" + sheet_name.replace("-", "_")
table = Table(displayName=table_name, ref=table_range)
table_style = TableStyleInfo(name="TableStyleMedium6", showRowStripes=True)
table.tableStyleInfo = table_style
ws.add_table(table)
Creates an Excel table for each sheet, applying a style and referencing the data range.

Step 5: Adding Summary Calculations
python
Copy
Edit
# Summary block on right
ws["I1"] = "Total Study Time"
ws["J1"] = "=SUMIF(D2:D100, \"Task\", C2:C100)/1440"
ws["I2"] = "Total Break Time"
ws["J2"] = "=SUMIF(D2:D100, \"Break\", C2:C100)/1440"
ws["I3"] = "Focus Ratio"
ws["J3"] = "=IF((J1+J2)=0, 0, J1/(J1+J2))"
Adds formulas for total study time, total break time, and focus ratio to the right side of each sheet. Focus ratio is calculated by dividing study time by the sum of study and break times.

Step 6: Creating the Summary Sheet
python
Copy
Edit
summary = wb.create_sheet("Summary")
summary_headers = ["Day", "Total Study Time", "Total Break Time", "Focus Ratio"]
Creates a new "Summary" sheet to aggregate data from all other sheets.

Adds headers for "Day", "Total Study Time", "Total Break Time", and "Focus Ratio".

python
Copy
Edit
for sheet in wb.sheetnames:
    if sheet == "Summary":
        continue
    sheet_ref = f"'{sheet}'"
    summary[f"A{row}"] = sheet
    summary[f"B{row}"] = f"=SUMIF({sheet_ref}!D2:D100, \"Task\", {sheet_ref}!C2:C100)/1440"
    summary[f"C{row}"] = f"=SUMIF({sheet_ref}!D2:D100, \"Break\", {sheet_ref}!C2:C100)/1440"
    summary[f"D{row}"] = f"=IF((B{row}+C{row})=0, 0, B{row}/(B{row}+C{row}))"
For each sheet (except "Summary"), it pulls the total study time, total break time, and focus ratio into the summary sheet using SUMIF formulas.

Step 7: Finalizing the Workbook
python
Copy
Edit
wb.save(file_path)
Saves the workbook to the specified file_path.

python
Copy
Edit
# Optional: For Colab download
from google.colab import files
files.download(file_path)
If running in Google Colab, this line allows you to download the generated Excel file directly to your local machine.

Conclusion
This script provides a comprehensive and automated time tracking log in Excel format, with built-in calculations and summaries, ideal for managing work/study sessions. It includes features such as automatic time entry, duration calculation, and a daily focus ratio, making it a powerful tool for tracking productivity. 

## License
MIT License

## Keywords:
time tracking, work log, study log, Excel automation, Python script
