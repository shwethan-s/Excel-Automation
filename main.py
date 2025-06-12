import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime

# Folder paths
PRE_UPDATED = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Pre-Updated"
CLEANED_OUTPUT = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Updated"

os.makedirs(CLEANED_OUTPUT, exist_ok=True)

def clean_and_convert_csv_to_excel(csv_path, output_excel_path):
    print(f"Cleaning CSV and converting to Excel: {csv_path}")
    try:
        # Load CSV
        df = pd.read_csv(csv_path)
        df.dropna(how='all', inplace=True)

        # Save as temporary Excel file
        df.to_excel(output_excel_path, index=False)

        # Format with openpyxl
        wb = load_workbook(output_excel_path)
        sheet = wb.active

        # Find first usable row
        min_row = 1
        for i in range(1, 6):
            values = [cell.value for cell in sheet[i] if cell.value is not None]
            if not values or all(str(v).strip() == "" for v in values):
                min_row = i + 1

        # Resize columns and wrap text
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row >= min_row:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                    cell.alignment = Alignment(wrap_text=True)
            sheet.column_dimensions[col_letter].width = max_len + 2

        # Removing the junk rows
        for _ in range(min_row - 1):
            sheet.delete_rows(1)

        wb.save(output_excel_path)
        print(f"Saved formatted Excel to: {output_excel_path}")

    except Exception as e:
        print(f"Error converting {csv_path}: {e}")

# Loop through and convert CSVs
for file in os.listdir(PRE_UPDATED):
    if file.lower().endswith(".csv"):
        timestamp = datetime.now().strftime("%m-%d-%Y_%I-%M%p")
        base_name = os.path.splitext(file)[0]
        output_file_name = f"{timestamp}_{base_name}.xlsx"

        full_input_path = os.path.join(PRE_UPDATED, file)
        full_output_path = os.path.join(CLEANED_OUTPUT, output_file_name)

        clean_and_convert_csv_to_excel(full_input_path, full_output_path)
        
