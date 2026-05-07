import pandas as pd

import numpy as np

from pathlib import Path

import os

import re

import json

import warnings
 
warnings.filterwarnings("ignore")
 
mapping_file = r"config/mapping.json"
 
def load_mapping():

    with open(mapping_file) as f:

        return json.load(f)
 
def reverse_mapping(mapping):

    return {v: k for k, v in mapping.items()}
 
def extract_rpt_id(filename):

    match = re.search(r"(RPT\d+[A-Z]{2})", filename.upper())

    return match.group(1) if match else None
 
def extract_tc(filename):

    match = re.search(r"TC(\d+)", filename.upper())

    return f"TC{match.group(1)}" if match else None
 
 
# -------------------- HEADER --------------------
 
def forward_fill_header(df):

    df_processed = df.copy()

    df_processed.columns = range(len(df_processed.columns))
 
    section_row = df_processed.iloc[0].fillna("")

    group_row = df_processed.iloc[2].fillna("")

    main_header_row = df_processed.iloc[3].fillna("")
 
    current_section = ""

    for i in range(len(section_row)):

        val = str(section_row.iloc[i])

        if val.strip() not in ("", "nan"):

            current_section = val

        section_row.iloc[i] = current_section
 
    current_group = ""

    for i in range(len(group_row)):

        val = str(group_row.iloc[i])

        if val.strip() not in ("", "nan"):

            current_group = val

        group_row.iloc[i] = current_group
 
    new_columns = []

    seen_columns = {}
 
    for i in range(len(main_header_row)):

        section = str(section_row.iloc[i]).strip()

        group = str(group_row.iloc[i]).strip()

        main_header = str(main_header_row.iloc[i]).strip()
 
        parts = [p for p in [section, group, main_header] if p and p != "nan"]
 
        if parts:

            col_name = "_".join(parts).replace("\n", " ").replace("\r", " ")

            col_name = " ".join(col_name.split())
 
            if col_name in seen_columns:

                seen_columns[col_name] += 1

                col_name = f"{col_name}_{seen_columns[col_name]}"

            else:

                seen_columns[col_name] = 1
 
            new_columns.append(col_name)

        else:

            new_columns.append(f"Column_{i}")
 
    df_data = df_processed.iloc[4:].copy()

    df_data.columns = new_columns

    df_data = df_data.dropna(how="all").reset_index(drop=True)
 
    return df_data
 
 
# -------------------- CLEANING --------------------
 
def clean_dataframe(df):
 
    df_clean = df.copy()
 
    # ---- NEW LOGIC: FIRST 7 COLUMNS AS KEY ----

    key_columns = list(df_clean.columns[:7])
 
    print(f"Using first 7 columns as key: {key_columns}")
 
    # Convert to string and clean

    for col in key_columns:

        df_clean[col] = df_clean[col].astype(str).str.strip()

        df_clean[col] = df_clean[col].replace(["nan", ""], np.nan)
 
    # Drop rows where ALL key columns are null

    df_clean = df_clean.dropna(subset=key_columns, how="all")
 
    # Create composite key

    df_clean["COMPOSITE_KEY"] = df_clean[key_columns].fillna("").agg("||".join, axis=1)
 
    primary_key = "COMPOSITE_KEY"
 
    # Clean remaining columns

    for col in df_clean.columns:

        if col != primary_key:

            try:

                df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce")

            except:

                df_clean[col] = df_clean[col].astype(str).str.strip()

                df_clean[col] = df_clean[col].replace(["nan", ""], np.nan)
 
    return df_clean, primary_key
 
 
# -------------------- COMPARISON --------------------
 
def compare_dataframes(df1, df2, primary_key, tolerance=0.0001):
 
    common_cols = list(set(df1.columns) & set(df2.columns))

    common_cols = [c for c in common_cols if c != primary_key]
 
    common_keys = set(df1[primary_key]) & set(df2[primary_key])
 
    mismatches = []
 
    print(f"Comparing {len(common_keys)} records...")
 
    for key in common_keys:
 
        row1 = df1[df1[primary_key] == key].iloc[0]

        row2 = df2[df2[primary_key] == key].iloc[0]
 
        for col in common_cols:
 
            v1 = row1[col]

            v2 = row2[col]
 
            if pd.isna(v1) and pd.isna(v2):

                continue
 
            if pd.isna(v1) or pd.isna(v2):

                mismatches.append({

                    "KEY": key,

                    "Column": col,

                    "Value_App": str(v1) if not pd.isna(v1) else "NULL",

                    "Value_Excel": str(v2) if not pd.isna(v2) else "NULL",

                })

            else:

                try:

                    if abs(float(v1) - float(v2)) > tolerance:

                        mismatches.append({

                            "KEY": key,

                            "Column": col,

                            "Value_App": float(v1),

                            "Value_Excel": float(v2),

                        })

                except:

                    if str(v1).strip() != str(v2).strip():

                        mismatches.append({

                            "KEY": key,

                            "Column": col,

                            "Value_App": str(v1),

                            "Value_Excel": str(v2),

                        })
 
    return mismatches
 
 
# -------------------- HTML --------------------
 
def generate_html_report(df_app, df_excel, mismatches, primary_key, file_names):
 
    common_keys = set(df_app[primary_key]) & set(df_excel[primary_key])

    unique_app = set(df_app[primary_key]) - set(df_excel[primary_key])

    unique_excel = set(df_excel[primary_key]) - set(df_app[primary_key])
 
    mismatches_by_column = {}

    for m in mismatches:

        mismatches_by_column[m["Column"]] = mismatches_by_column.get(m["Column"], 0) + 1
 
    html = f"""
<html>
<head>
<title>Comparison Report</title>
<style>

body {{ font-family: Arial; }}

table {{ border-collapse: collapse; width: 100%; }}

th, td {{ border: 1px solid black; padding: 5px; }}

th {{ background-color: #f2f2f2; }}
</style>
</head>
<body>
<h2>Summary</h2>
<table>
<tr><td>App Rows</td><td>{len(df_app)}</td></tr>
<tr><td>Excel Rows</td><td>{len(df_excel)}</td></tr>
<tr><td>Common Records</td><td>{len(common_keys)}</td></tr>
<tr><td>Unique App</td><td>{len(unique_app)}</td></tr>
<tr><td>Unique Excel</td><td>{len(unique_excel)}</td></tr>
<tr><td>Total Mismatches</td><td>{len(mismatches)}</td></tr>
</table>
 
<h2>Mismatch by Column</h2>
<table>
<tr><th>Column</th><th>Count</th></tr>

"""
 
    for col, cnt in sorted(mismatches_by_column.items(), key=lambda x: x[1], reverse=True):

        html += f"<tr><td>{col}</td><td>{cnt}</td></tr>"
 
    html += """
</table>
 
<h2>Mismatch Details</h2>
<table>
<tr><th>#</th><th>Key</th><th>Column</th><th>App</th><th>Excel</th><th>Diff</th></tr>

"""
 
    for i, m in enumerate(mismatches, 1):

        diff = ""

        try:

            diff = abs(float(m["Value_App"]) - float(m["Value_Excel"]))

        except:

            pass
 
        html += f"""
<tr>
<td>{i}</td>
<td>{m['KEY']}</td>
<td>{m['Column']}</td>
<td>{m['Value_App']}</td>
<td>{m['Value_Excel']}</td>
<td>{diff}</td>
</tr>

"""
 
    html += "</table></body></html>"
 
    return html
 
 
# -------------------- MAIN --------------------
 
def main():
 
    base_path = Path("data/finAlign")

    file1_path = base_path/"File1"

    file2_path = base_path/"File2"

    output_path = base_path/"Output"
 
    os.makedirs(output_path, exist_ok=True)
 
    mapping = load_mapping()

    rev_map = reverse_mapping(mapping)
 
    file1_dict = {}

    file2_dict = {}
 
    for f in file1_path.glob("*.xlsx"):

        rpt = extract_rpt_id(f.name)

        if rpt and rpt in rev_map:

            tc = rev_map[rpt]

            file1_dict[tc] = f
 
    for f in file2_path.glob("*.xlsx"):

        tc = extract_tc(f.name)

        if tc:

            file2_dict[tc] = f
 
    common_tcs = set(file1_dict) & set(file2_dict)
 
    print(f"Running {len(common_tcs)} scenarios...")
 
    for tc in sorted(common_tcs):
 
        print(f"\nProcessing {tc}...")
 
        df1 = pd.read_excel(file1_dict[tc], header=None)

        df2 = pd.read_excel(file2_dict[tc], header=None)
 
        df1 = forward_fill_header(df1)

        df2 = forward_fill_header(df2)
 
        df1, pk = clean_dataframe(df1)

        df2, _ = clean_dataframe(df2)
 
        mismatches = compare_dataframes(df1, df2, pk)
 
        output_file = output_path / f"{tc}_FinAlign_Dev6May.html"
 
        html = generate_html_report(df1, df2, mismatches, pk, output_file)
 
        with open(output_file, "w", encoding="utf-8") as f:

            f.write(html)
 
    print("All scenarios completed.")
 
 
if __name__ == "__main__":

    main()