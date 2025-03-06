import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def merge_sheets(files):
    data = [pd.read_excel(f, sheet_name=None, dtype=str) for f in files]
    sheet_names = data[0].keys()
    merged_sheets = {}

    for sheet_name in sheet_names:
        dfs = [d[sheet_name] for d in data]
        merged_df = None

        for df in dfs:
            df.dropna(how='all', inplace=True)
            header_row = df[df.applymap(lambda x: str(x).strip().lower()).astype(str)
                             .apply(lambda x: x.str.contains('total', case=False, na=False)).any(axis=1)].index.min()
            start_row = header_row + 1 if pd.notna(header_row) else 0
        
        dfs = [df.iloc[start_row:].reset_index(drop=True) for df in dfs]
        max_rows = max(df.shape[0] for df in dfs)
        max_cols = max(df.shape[1] for df in dfs)
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
        merged_df.dropna(how='all', axis=0, inplace=True)
        merged_df.dropna(how='all', axis=1, inplace=True)
        merged_sheets[sheet_name] = merged_df
    
    return merged_sheets

st.title("Excel Sheet Merger")

uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    merged_sheets = merge_sheets(uploaded_files)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in merged_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    output.seek(0)
    
    st.download_button("Download Merged Excel File", data=output, file_name="merged_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
