# Excel Automation Tool

A Python-based tool that automates the cleaning and consolidation of messy monthly Excel files into a master workbook. Designed to be simple enough for non-technical users, with a GUI-based interface.

## 📌 Features

- Automatically removes top banners, logos, and unnecessary headers
- Unmerges cells and formats sheets into clean, structured tables
- Deletes empty columns and rows
- Computes and compares balances across monthly files
- Flags abnormal readings compared to previous years
- Exports a master workbook with consolidated, cleaned data
- Easy-to-use Windows GUI — no command-line knowledge required

## 🛠 Technologies Used

- Python 3.13.3
- `pandas` for data manipulation
- `openpyxl` for Excel read/write
- `tkinter` (or `PyQt` / `customtkinter`) for GUI
- `pyinstaller` for building into a Windows app


