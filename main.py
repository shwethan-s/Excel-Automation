import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# Folder paths
PRE_UPDATED = "C:\Users\Shwethan\OneDrive\Shwethan - McMaster University\Billing\Pre-Updated"
CLEANED_OUTPUT = "C:\Users\Shwethan\OneDrive\Shwethan - McMaster University\Billing\Updated"

os.makedirs(CLEANED_OUTPUT, exist_ok=True)

def clean_excel_file(filepath, output_path):
    print(f"🧹 Cleaning: {filepath}")
    
    wb = load_workbook(filepath)
    sheet = wb.active

    # Step 1: Unmerge all cells
    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))

    # Step 2: Remove top junk rows (assume useful data starts below row 5)
    min_row = 1
    for i in range(1, 6):
        values = [cell.value for cell in sheet[i] if cell.value is not None]
        if not values or all(str(v).strip() == "" for v in values):
            min_row = i + 1

    # Step 3: Adjust column widths and apply word wrapping
    for col in sheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= min_row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
                cell.alignment = Alignment(wrap_text=True)
        sheet.column_dimensions[col_letter].width = max_len + 2

    # Step 4: Remove rows above min_row
    for i in range(min_row - 1):
        sheet.delete_rows(1)

    # Step 5: Save the cleaned file
    wb.save(output_path)
    print(f"✅ Saved cleaned file to: {output_path}")

# Loop through Excel files and clean them
for file in os.listdir(PRE_UPDATED):
    if file.endswith(".xlsx"):
        full_path = os.path.join(PRE_UPDATED, file)
        output_file = os.path.join(CLEANED_OUTPUT, f"cleaned_{file}")
        clean_excel_file(full_path, output_file)
