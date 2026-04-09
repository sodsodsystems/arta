import pandas as pd

file_path = r'c:\Users\admin\Documents\GIT PROJECTS\ARTA\SRS HR COPY.xlsx'
try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheet names: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        print(f"\n--- {sheet} ---")
        df = pd.read_excel(file_path, sheet_name=sheet).head(10)
        print(df.to_string())
except Exception as e:
    print(f"Error: {e}")
