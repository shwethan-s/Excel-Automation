import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime
import shutil

# Paths
PRE_UPDATED = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Pre-Updated"
INTERMEDIATE_FOLDER = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Intermediate Folder"
OUTPUT_FOLDER = "C:\\Users\\sivaps15\\OneDrive - McMaster University\\Billing\\Output"


os.makedirs(INTERMEDIATE_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def format_excel(input_path, intermediate_subfolder, master_data, today_str, bill_month):
    filename = os.path.basename(input_path)
    temp_path = input_path

    if input_path.lower().endswith(".xls"):
        df = pd.read_excel(input_path, engine="xlrd")
        temp_path = os.path.join(intermediate_subfolder, filename.replace(".xls", "_temp.xlsx"))
        df.to_excel(temp_path, index=False)

    # Load workbook
    wb = load_workbook(temp_path)
    sheet = wb.active

    # Find "Timestamp"
    timestamp_row = None
    for row in sheet.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and "timestamp" in str(cell.value).lower():
                timestamp_row = cell.row
                break
        if timestamp_row:
            break

    if timestamp_row is None:
        print(f"⚠️ Timestamp not found in: {filename}")
        return

    # Remove all rows above timestamp, but leave 3 rows above it
    if timestamp_row > 4:
        for _ in range(timestamp_row - 4):
            sheet.delete_rows(1)

    # Remove empty columns
    valid_cols = []
    for col in sheet.iter_cols(min_row=4, max_col=sheet.max_column):
        if any(cell.value not in (None, "") for cell in col):
            valid_cols.append(col[0].column)

    for col_idx in reversed(range(1, sheet.max_column + 1)):
        if col_idx not in valid_cols:
            sheet.delete_cols(col_idx)

    # Add usage calculation in 3 rows above timestamp
    first_row = 4
    for col in sheet.iter_cols(min_row=first_row + 1, min_col=2):
        col_values = [cell.value for cell in col if isinstance(cell.value, (int, float))]

        if len(col_values) < 2:
            continue

        usage = col_values[-1] - col_values[0]
        meter_name = col[0].column_letter

        # Timestamp check
        time_col = sheet["A"]
        try:
            first_time = str(time_col[first_row].value).strip()
            last_time = str(time_col[first_row + len(col_values) - 1].value).strip()
        except:
            first_time, last_time = "", ""

        correct = ("12:15 AM" in first_time and "1" in first_time) and ("12:00 AM" in last_time and "1" in last_time)
        color = "C6EFCE" if correct else "FFC7CE"  # Green or red

        usage_cell = sheet.cell(row=1, column=col[0].column)
        usage_cell.value = round(usage, 2)
        usage_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        usage_cell.alignment = Alignment(horizontal="center")

        master_data.append([intermediate_subfolder.split("_")[-1], meter_name, round(usage, 2)])

    # Format as table
    end_col = get_column_letter(sheet.max_column)
    end_row = sheet.max_row
    table = Table(displayName="MeterTable", ref=f"A{first_row}:{end_col}{end_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    table.tableStyleInfo = style
    sheet.add_table(table)

    # Resize columns and align text
    for col in sheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        for cell in col:
            cell.alignment = Alignment(wrap_text=True)
        sheet.column_dimensions[col_letter].width = max_len + 2

    # Save final
    output_path = os.path.join(intermediate_subfolder, filename.replace(".xls", ".xlsx"))
    wb.save(output_path)

    if temp_path != input_path:
        os.remove(temp_path)

def main():
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")
    bill_month = today.strftime("%B")
    master_data = []

    for file in os.listdir(PRE_UPDATED):
        if file.endswith((".xls", ".xlsx")):
            building = os.path.splitext(file)[0].split("Residence - ")[-1].split(" Bld")[0].strip()
            subfolder_name = f"{today_str}_{bill_month}_{building}"
            intermediate_subfolder = os.path.join(INTERMEDIATE_FOLDER, subfolder_name)
            os.makedirs(intermediate_subfolder, exist_ok=True)

            file_path = os.path.join(PRE_UPDATED, file)
            format_excel(file_path, intermediate_subfolder, master_data, today_str, bill_month)

    # Save master output
    if master_data:
        master_df = pd.DataFrame(master_data, columns=["Building", "Meter", "Usage"])
        master_filename = f"Final-{today_str}-{bill_month}.xlsx"
        master_df.to_excel(os.path.join(OUTPUT_FOLDER, master_filename), index=False)

if __name__ == "__main__":
    main()
