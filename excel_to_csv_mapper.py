import openpyxl
import csv
import os

# --------- Step 1: Load Excel ----------
excel_file = "./excel_csv_mapper/WPT_JR(B)26B1.xlsx"
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active  # take the first sheet

# --------- Step 2: CSV file ----------
csv_file = "./excel_csv_mapper/WPT_JR(B)26B1_csv - Copy.csv"

# Load existing CSV data if exists
csv_data = []
if os.path.exists(csv_file):
    with open(csv_file, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        csv_data = list(reader)

# --------- Step 3: User mapping ----------
# User defines mappings: (source_cell, target_column_index)
# target_column_index starts at 0
mappings = [
    ("B4", 4),  # C3 → column B (index 1)
    ("C4", 7),  # D4 → column A (index 0)
    ("H4", 3),
    ("I4", 0),
    ("L4", 8),
    ("M4", 6),
]

# --------- Step 4: Get data from Excel ----------
new_row = []
for _, target_col in mappings:
    # ensure row is big enough
    while len(new_row) <= target_col:
        new_row.append("")
for source_cell, target_col in mappings:
    value = sheet[source_cell].value
    new_row[target_col] = value

# --------- Step 5: Append to CSV ----------
csv_data.append(new_row)

# --------- Step 6: Save CSV ----------
with open(csv_file, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerows(csv_data)

print("✅ Data transferred successfully!")