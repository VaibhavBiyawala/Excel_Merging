import re
import numpy as np
import streamlit as st
import pandas as pd
from io import BytesIO

def merge_files(files):
    data = [pd.read_excel(f, sheet_name=None, dtype=str) for f in files]
    sheet_names = list(data[0].keys())[1:]  # Skip first sheet

    special_sheets = ["StaffSalaries", "Nursery", "LKG", "Pre-First"]
    for i in range(1, 13):
        suffix = "th" if i > 3 else ("st" if i == 1 else ("nd" if i == 2 else "rd"))
        special_sheets.append(f"{i}{suffix}")

    merged_sheets = {}

    for sheet_name in sheet_names:
        if sheet_name in special_sheets:
            dfs = [d[sheet_name] for d in data]
            headers, section_headers, section_data = [], [], {}
            first_section_row = None
            
            for i, row in dfs[0].iterrows():
                row_values = [str(x).strip() for x in row.values]
                first_cell = row_values[0] if row_values else ""
                if re.match(r'^\d+\.', first_cell):
                    first_section_row = i
                    break
                else:
                    headers.append(row.values)
            
            if first_section_row is None:
                first_section_row = 0
            
            for df in dfs:
                current_section = None
                for i, row in df.iterrows():
                    if i < first_section_row:
                        continue
                    
                    row_values = [str(x).strip() for x in row.values]
                    first_cell = row_values[0] if row_values else ""
                    
                    if re.match(r'^\d+\.', first_cell):
                        current_section = first_cell
                        if first_cell not in section_data:
                            section_headers.append((first_cell, int(first_cell.split('.')[0])))
                            section_data[first_cell] = {"header": row.values, "content": []}
                    elif current_section:
                        lower_values = [str(x).lower().strip() for x in row_values]
                        if any("total" in val for val in lower_values if isinstance(val, str)):
                            continue
                        if all(not val or val in ["0", "-", "nan"] for val in lower_values):
                            continue
                        section_data[current_section]["content"].append(row.values)
            
            section_headers.sort(key=lambda x: x[1])
            result_rows = []
            for header in headers:
                result_rows.append(header)
            
            rate_col_idx = None
            for header_row in headers:
                for col_idx, cell_value in enumerate(header_row):
                    if str(cell_value).strip().lower() == "rate":
                        rate_col_idx = col_idx
                        break
                if rate_col_idx is not None:
                    break
            
            for section_name, _ in section_headers:
                result_rows.append(section_data[section_name]["header"])
                content_rows = section_data[section_name]["content"]
                for row in content_rows:
                    result_rows.append(row)
                
                if content_rows:
                    numeric_totals = []
                    for col_idx in range(2, len(content_rows[0])):
                        if rate_col_idx is not None and col_idx == rate_col_idx:
                            numeric_totals.append(0)
                            continue
                            
                        col_total = 0
                        for row in content_rows:
                            if col_idx < len(row):
                                if rate_col_idx is not None and rate_col_idx < len(row):
                                    qty_value = str(row[col_idx]).strip()
                                    rate_value = str(row[rate_col_idx]).strip()
                                    
                                    if qty_value and qty_value not in ["-", "nan"] and rate_value and rate_value not in ["-", "nan"]:
                                        try:
                                            col_total += float(qty_value.replace(',', '')) * float(rate_value.replace(',', ''))
                                        except ValueError:
                                            pass
                                else:
                                    cell_value = str(row[col_idx]).strip()
                                    if cell_value and cell_value not in ["-", "nan"]:
                                        try:
                                            col_total += float(cell_value.replace(',', ''))
                                        except ValueError:
                                            pass
                        numeric_totals.append(col_total)
                    
                    subtotal_row = ["Sub-Total"] + [""] * (len(content_rows[0]) - 1)
                    if any(val != 0 for val in numeric_totals):
                        for idx, val in enumerate(numeric_totals):
                            subtotal_row[idx + 2] = f"{int(val):,}" if val == int(val) else f"{val:,.2f}"
                        result_rows.append(subtotal_row)
            
            merged_df = pd.DataFrame(result_rows)
            for c in range(2, merged_df.shape[1]):
                merged_df[c] = pd.to_numeric(merged_df[c].astype(str).str.replace(',', ''), errors='ignore')
            
            merged_df.replace([0, "-"], np.nan, inplace=True)
            merged_df = merged_df.dropna(how='all', axis=0).dropna(how='all', axis=1)
            
        else:
            dfs = [d[sheet_name] for d in data]
            max_rows = max(df.shape[0] for df in dfs)
            max_cols = max(df.shape[1] for df in dfs)
            merged_df = pd.DataFrame(index=range(max_rows), columns=range(max_cols))
            
            for r in range(max_rows):
                for c in range(max_cols):
                    values = []
                    for df in dfs:
                        try:
                            val = df.iloc[r, c]
                            if not pd.isna(val) and val != "-":
                                values.append(val)
                        except:
                            pass
                    
                    if values and all(str(v).replace('.', '', 1).replace(',', '').isdigit() for v in values if pd.notna(v) and v != ""):
                        converted_values = []
                        for v in values:
                            try:
                                if isinstance(v, str):
                                    v = v.replace(',', '')
                                converted_values.append(float(v))
                            except (ValueError, TypeError):
                                pass
                        
                        merged_df.iloc[r, c] = sum(converted_values) if converted_values else ""
                    elif values:
                        merged_df.iloc[r, c] = values[0]
                    else:
                        merged_df.iloc[r, c] = ""
            
            for c in range(merged_df.shape[1]):
                merged_df[c] = merged_df[c].astype(str).str.strip()
                merged_df[c] = merged_df[c].replace(['nan', 0, 0.0, '0', '-', ''], np.nan)
            
            merged_df = merged_df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        
        merged_sheets[sheet_name] = merged_df

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet, df in merged_sheets.items():
            df = df.astype(object)
            for col in df.columns:
                df[col] = df[col].astype(object).replace(['nan', 0, '0.0', '0', ''], np.nan).astype(object).fillna('')
            
            non_empty_rows = ~df.astype(str).apply(lambda x: x.str.strip().eq('').all(), axis=1)
            df[non_empty_rows].to_excel(writer, sheet_name=sheet, index=False, header=False)
    
    output.seek(0)
    return output

st.title("Excel Merge Tool")

uploaded_files = st.file_uploader(
    "Select one or more Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("Merge Files"):
    if not uploaded_files:
        st.warning("Please upload at least one Excel file.")
    else:
        st.write("Merging...")
        merged_file = merge_files(uploaded_files)
        st.success("Merging complete!")
        st.download_button(
            label="Download Merged File",
            data=merged_file,
            file_name="merged.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )