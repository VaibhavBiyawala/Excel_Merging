import pandas as pd
import numpy as np

files = [
    "Pre_Primary.xlsx",
    "PRIMARY.xlsx",  
]

# Read all sheets from all files
data = [pd.read_excel(f, sheet_name=None, dtype=str) for f in files]

# Get all sheet names
sheet_names = data[0].keys()

# Dictionary to store merged sheets
merged_sheets = {}

for sheet_name in sheet_names:
    # Identify the first row with actual data
    dfs = [d[sheet_name] for d in data]
    
    merged_df = None

    for df in dfs:
        df.dropna(how='all', inplace=True)  # Remove empty rows
        header_row = df[df.applymap(lambda x: str(x).strip().lower()).astype(str).apply(lambda x: x.str.contains('total', case=False, na=False)).any(axis=1)].index.min()
        start_row = header_row + 1 if pd.notna(header_row) else 0
    
    # Read sheets again with detected header row
    dfs = [df.iloc[start_row:].reset_index(drop=True) for df in dfs]
    
    # Find max row and column size to ensure all data is covered
    max_rows = max(df.shape[0] for df in dfs)
    max_cols = max(df.shape[1] for df in dfs)
    
    # Create an empty DataFrame to store merged data
    merged_df = pd.DataFrame(index=range(max_rows), columns=range(max_cols))
    
    for r in range(max_rows):
        for c in range(max_cols):
            values = []
            for df in dfs:
                try:
                    val = df.iloc[r, c]
                    if pd.isna(val) or val == "-":
                        continue
                    values.append(val)
                except:
                    pass  
            
            if all(v.replace('.', '', 1).isdigit() for v in values if isinstance(v, str)):
                merged_df.iloc[r, c] = sum(float(v) for v in values)
            elif values:
                merged_df.iloc[r, c] = values[0]  
            else:
                merged_df.iloc[r, c] = ""
    
    merged_df.replace(0, np.nan, inplace=True)
    merged_df.replace("-", np.nan, inplace=True)
    merged_df = merged_df.dropna(how='all', axis=0)  # Drop empty rows
    merged_df = merged_df.dropna(how='all', axis=1)  # Drop empty columns
    merged_sheets[sheet_name] = merged_df

# Save to a new Excel file
output_file = "result5.xlsx"
with pd.ExcelWriter(output_file) as writer:
    for sheet, df in merged_sheets.items():
        df.to_excel(writer, sheet_name=sheet, index=False, header=False)

print(f"Merged file saved to: {output_file}")