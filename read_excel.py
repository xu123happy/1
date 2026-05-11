import pandas as pd
import sys

file_path = 'C:/Users/91340/Desktop/临时文件/日报 - 副本.xlsx'
print(f"Reading: {file_path}")
try:
    xf = pd.ExcelFile(file_path)
    print('Sheets:', xf.sheet_names)
    for sheet in xf.sheet_names[:3]:
        print(f'\n=== Sheet: {sheet} ===')
        df = pd.read_excel(file_path, sheet_name=sheet, nrows=5)
        print('Columns:', list(df.columns))
        print(df.head(3).to_string())
except Exception as e:
    print(f'Error: {e}', file=sys.stderr)
    import traceback
    traceback.print_exc()
