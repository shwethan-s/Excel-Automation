import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime

# Required if handling .xls files
# pip install xlrd openpyxl

# Input and output paths
PRE_UPDATED = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Pre-Updated"
CLEANED_OUTPUT = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Updated"

os.makedirs(CLEANED_OUTPUT, exist_ok=True)

def clean_and_format_excel(input_path, output_path):
    # Step 1: Convert to .xlsx (if .xls)
    temp_path = output_path
    if input_path.lower().endswith(".xls"):
        df = pd.read_excel(input_path, engine="xlrd")
        temp_path = output_path.replace(".xlsx", "_temp.xlsx")
        df.to_excel(temp_path, index=False)

    # Step 2: Load workbook with openpyxl
    wb = load_workbook(temp_path)
    sheet = wb.active

    # Step 3: Find "Timestamp" row
    timestamp_row = None
    for row in sheet.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and "timestamp" in str(cell.value).lower():
                timestamp_row = cell.row
                break
        if timestamp_row:
            break

    # Step 4: Delete rows above timestamp row, leave 3 blank rows
    if timestamp_row and timestamp_row > 4:
        for _ in range(timestamp_row - 4):
            sheet.delete_rows(1)

    # Step 5: Delete empty columns from timestamp row onwards
    valid_cols = []
    for col in sheet.iter_cols(min_row=4, max_col=sheet.max_column):
        if any(cell.value not in (None, "") for cell in col):
            valid_cols.append(col[0].column)

    all_cols = [col[0].column for col in sheet.iter_cols(min_row=1, max_row=1)]
    for col_idx in reversed(all_cols):
        if col_idx not in valid_cols:
            sheet.delete_cols(col_idx)

    # Step 6: Resize and wrap text
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
            cell.alignment = Alignment(wrap_text=True)
        sheet.column_dimensions[col_letter].width = max_len + 2

    # Step 7: Save final result
    wb.save(output_path)

    # Delete temp file if needed
    if temp_path != output_path and os.path.exists(temp_path):
        os.remove(temp_path)

# Loop through files
for file in os.listdir(PRE_UPDATED):
    if file.lower().endswith((".xls", ".xlsx")):
        base_name = os.path.splitext(file)[0]
        timestamp = datetime.now().strftime("%m-%d-%Y_%I-%M%p")
        output_file = f"{timestamp}_{base_name}.xlsx"

        input_file_path = os.path.join(PRE_UPDATED, file)
        output_file_path = os.path.join(CLEANED_OUTPUT, output_file)

        clean_and_format_excel(input_file_path, output_file_path)

        # Remove original
        os.remove(input_file_path)
