📘 UPSC Study Log – June 2025
This Python script automatically generates a fully formatted Excel-based daily study tracker for the entire month of June 2025, designed to help UPSC aspirants plan, log, and analyze their daily preparation with precision.

🔑 Key Features:
📅 One Sheet per Day
Creates 30 separate sheets (June 1–30), each representing a day with pre-defined columns for time logging and task tracking.

🕒 Smart Time Logging

Start Time, End Time, and auto-calculated Duration in minutes

Auto-fill formulas chain Start Times from the previous End Time

🔽 Drop-Down for Task Type
Each row includes a drop-down to mark the entry as either a Task or a Break

🧮 Daily Summary Section
Each day's sheet calculates:

Total Study Time

Total Break Time

Focus Ratio (Study / Total Time)

🎨 Clean Visual Formatting

Bold headers, colored columns, and styled tables

Column widths auto-adjusted for readability

Sheet tabs colored for easy navigation

📊 Monthly Summary Sheet

Aggregates daily totals and focus ratios

Displays the data in a structured summary table for quick review

🏠 Home Sheet with Hyperlinks

Central landing page with clickable links to all 30 daily sheets

Improves navigation across the workbook

📥 Ready for Download (Google Colab Compatible)
Automatically triggers download after generation if used in Google Colab

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Code explanation

📦 Library Imports
python
Copy
Edit
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from datetime import datetime
pandas is used to initially create and write blank DataFrames to Excel sheets.

openpyxl handles all the formatting, formulas, styling, and worksheet manipulation.

Styling tools like Font, PatternFill, and Alignment allow detailed customization.

Table, TableStyleInfo — define Excel tables with stylings.

DataValidation — used for dropdowns.

get_column_letter helps with Excel-style column referencing (like A, B, C…).

datetime is used to define the study period (June 2025).

🗓️ Configuration Section
python
Copy
Edit
start_date = datetime(2025, 6, 1)
end_date = datetime(2025, 6, 30)
date_range = pd.date_range(start=start_date, end=end_date)
Define the month range — here June 2025.

date_range contains all 30 dates from June 1 to June 30.

python
Copy
Edit
columns = [
    "Start Time", "End Time", "Duration (minutes)",
    "Type", "Task Name", "Description / Subject", "Remarks"
]
Defines column headers for each sheet (one per day).

These represent logged time and tasks.

python
Copy
Edit
header_colors = [
    "FBE5D6", "D9EAD3", "D9D2E9", "CFE2F3", "FFF2CC", "F4CCCC", "EAD1DC"
]
These hex colors are for each column header to differentiate them visually.

python
Copy
Edit
header_fills = [PatternFill(start_color=c, end_color=c, fill_type="solid") for c in header_colors]
Convert colors into Excel fill styles.

python
Copy
Edit
header_font = Font(bold=True, size=16)
cell_font = Font(size=14)
wrap_alignment = Alignment(wrap_text=True, vertical="top")
summary_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
Define font and alignment styles for headers and cells.

summary_fill is used in the summary block for totals.

python
Copy
Edit
tab_colors = [
    ...
]
Colors for each day-tab. Repeats if more days than colors.

python
Copy
Edit
file_path = "/content/UPSC_Study_Log_June_2025.xlsx"
Output file path (for Google Colab).

📄 Step 1: Create Sheets with Placeholder Data
python
Copy
Edit
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    for i, date in enumerate(date_range, 1):
        sheet_name = str(i)
        df = pd.DataFrame([[""] * len(columns)], columns=columns)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
Creates 30 sheets named 1 to 30, each with a blank row and the predefined columns.

🎨 Step 2: Format Each Daily Sheet
python
Copy
Edit
wb = openpyxl.load_workbook(file_path)
Re-opens the Excel file for further styling and manipulation.

python
Copy
Edit
for i, sheet_name in enumerate(wb.sheetnames):
    if sheet_name == "Summary" or sheet_name == "Home":
        continue
Skip formatting Summary and Home (they’ll be created later).

🔹 Date Title
python
Copy
Edit
actual_date = date_range[i].strftime("%B %d, %Y")
...
ws.insert_rows(1)
ws["A1"] = actual_date
...
Adds the actual date (like “June 01, 2025”) in the top merged cell.

🔹 Header Styling
python
Copy
Edit
for col_idx, (fill, header) in enumerate(zip(header_fills, columns), 1):
    ...
Applies background fill and bold text to headers.

🔹 Row Formatting & Time Columns
python
Copy
Edit
for row in ws.iter_rows(min_row=3, max_row=102, min_col=1, max_col=len(columns)):
    ...
Formats cell fonts and text wrapping for the first 100 data rows (row 3 to 102).

python
Copy
Edit
ws[f"A{row}"].number_format = "h:mm:ss AM/PM"
...
ws[f"A{row}"] = f"=B{row - 1}"
...
Applies time formatting and auto-fill formulas:

Start time = Previous row’s End time

Duration = difference between End and Start

🔽 Dropdown for “Type”
python
Copy
Edit
dv = DataValidation(type="list", formula1='"Task,Break"', allow_blank=True)
ws.add_data_validation(dv)
dv.add("D3:D102")
Adds a dropdown to the “Type” column for selecting either Task or Break.

🔢 Add Table for Excel UI
python
Copy
Edit
table = Table(displayName=table_name, ref=table_range)
...
ws.add_table(table)
Adds an Excel table so the range looks neat and functions like a database.

📊 Summary Block (Daily)
python
Copy
Edit
ws["I2"] = "Total Study Time"
...
ws["J2"] = "=SUMIF(D3:D102, \"Task\", C3:C102)/1440"
Creates summary cells on the side:

Total Study Time → Sums all “Task” durations

Total Break Time → Sums all “Break” durations

Focus Ratio → Task / (Task + Break)

📐 Auto Adjust Column Width
python
Copy
Edit
for col_idx in range(1, len(columns) + 1):
    ...
Sets column width dynamically based on longest value in each.

🟦 Sheet Tab Colors
python
Copy
Edit
ws.sheet_properties.tabColor = tab_colors[i % len(tab_colors)]
Adds color to the bottom tab for visual navigation.

🧾 Step 3: Create Summary Sheet
python
Copy
Edit
summary = wb.create_sheet("Summary")
headers = ["Day", "Total Study Time", "Total Break Time", "Focus Ratio"]
...
Adds a new sheet summarizing all 30 days.

python
Copy
Edit
for i in range(1, len(date_range) + 1):
    ...
    summary[f"B{i+1}"] = f"=SUMIF('{i}'!D3:D102, \"Task\", '{i}'!C3:C102)/1440"
For each day, reference the matching sheet and summarize its totals.

python
Copy
Edit
summary_table = Table(displayName="SummaryTable", ref=summary_range)
summary.add_table(summary_table)
Adds Excel table formatting to the summary sheet.

🏠 Step 4: Create Home Sheet with Hyperlinks
python
Copy
Edit
home = wb.create_sheet("Home", 0)
...
for i in range(1, len(date_range) + 1):
    cell = home.cell(row=i + 1, column=1, value=f"Day {i}")
    ...
    cell.hyperlink = f"#{i}!A1"
Adds a front-page sheet named "Home" with clickable hyperlinks to each daily log.

💾 Final Save and Optional Download
python
Copy
Edit
wb.save(file_path)
Saves the workbook with all edits.

python
Copy
Edit
from google.colab import files
files.download(file_path)
For Colab: triggers download of the generated Excel file.

