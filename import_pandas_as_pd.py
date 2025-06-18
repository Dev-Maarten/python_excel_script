import pandas as pd
import os

# Find first Excel file in current directory
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]
if not excel_files:
    raise FileNotFoundError("No Excel files found in current directory")

file_path = excel_files[0]

xls = pd.ExcelFile(file_path)
print(f"File: {file_path}")
print("Available sheets:", xls.sheet_names)