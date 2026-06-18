import pandas as pd
try:
    df = pd.read_excel('manager_incharge_lookupss.xlsx', sheet_name=None)
    with open('headers.txt', 'w') as f:
        for sheet_name, sheet_df in df.items():
            f.write(f"Sheet: {sheet_name}\n")
            f.write(f"Headers: {sheet_df.columns.tolist()}\n")
            f.write(f"First row: {sheet_df.iloc[0].tolist() if not sheet_df.empty else 'empty'}\n\n")
except Exception as e:
    with open('headers.txt', 'w') as f:
        f.write(f"Error: {e}")
