import os
import pandas as pd
from openpyxl import load_workbook

# Folder containing messy monthly sheets
INPUT_FOLDER = "C:\Users\sivaps15\OneDrive - McMaster University\Billing\Pre-Updated"
OUTPUT_FILE = "../Output/MasterWorkbook.xlsx"

def clean_excel_file(filepath):
    """Load them onto a clean file, drop empty rows/columns, and fill NaNs."""
    print(f"Processing: {filepath}")
    try:
        df = pd.read_excel(filepath, engine='openpyxl')

        # Drop empty columns and rows
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)
        df.dropna(how='any', axis=1, inplace=True)  # Drop columns with any NaN values

        # Fill NaNs or fix formatting issues as needed
        df = df.fillna("")

        return df

    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return pd.DataFrame()

def process_all_excels():
    print("Starting Excel cleaning and consolidation...")

    combined_data = []

    for filename in os.listdir(INPUT_FOLDER):
        if filename.endswith(".xlsx"):
            path = os.path.join(INPUT_FOLDER, filename)
            df = clean_excel_file(path)
            if not df.empty:
                df["Source File"] = filename  # Add origin info
                combined_data.append(df)

    if not combined_data:
        print("No valid Excel files found.")
        return

    master_df = pd.concat(combined_data, ignore_index=True)
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    master_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Master workbook saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    process_all_excels()


