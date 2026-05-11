import pandas as pd
import sys

file_path = r'C:\Users\91340\Desktop\临时文件\日报 - 副本.xlsx'
print("Reading:", file_path)

try:
    xf = pd.ExcelFile(file_path)
    print("Sheets found:", xf.sheet_names)
    
    for sheet in xf.sheet_names:
        print(f"\n=== Sheet: {sheet} ===")
        df = pd.read_excel(file_path, sheet_name=sheet)
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print("First 5 rows:")
        print(df.head(10).to_string())
        print()
except Exception as e:
    print("Error:", str(e))
    import traceback
    traceback.print_exc()
