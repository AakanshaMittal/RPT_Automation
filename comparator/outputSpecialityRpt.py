import pandas as pd
import numpy as np
import os
import json
import re
from pathlib import Path
from collections import defaultdict
from html import escape
 
APP_FOLDER = r"data/outputSpecialityRpt/File1"
EXCEL_FOLDER = r"data/outputSpecialityRpt/File2"
OUTPUT_FOLDER = r"data/outputSpecialityRpt/Output"
mapping_file = r"config/mapping.json"
 
 
HEADER_ROW_1 = 12

HEADER_ROW_2 = 13

NDC_COL_NAME = "NDC"

TOLERANCE = 0.001
 
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
    return match.group(1) if match else None 
 
def load_and_prepare(file_path):

    df_raw = pd.read_excel(file_path, header=None)
 
    # Ignore D13

    df_raw.iat[HEADER_ROW_1, 3] = np.nan
 
    header1 = df_raw.iloc[HEADER_ROW_1].ffill()

    header1.iloc[3] = ""
 
    header2 = df_raw.iloc[HEADER_ROW_2]
 
    combined_headers = []
 
    for h1, h2 in zip(header1, header2):

        h1 = str(h1).strip() if pd.notna(h1) else ""

        h2 = str(h2).strip() if pd.notna(h2) else ""
 
        if h1 and h2:

            combined_headers.append(f"{h1}_{h2}")

        elif h2:

            combined_headers.append(h2)

        else:

            combined_headers.append(h1)
 
    df = df_raw.iloc[HEADER_ROW_2 + 1:].copy()

    df.columns = combined_headers

    df = df.dropna(how="all")
 
    df.columns = [normalize_text(c) for c in df.columns]
 
    ndc_col = find_ndc_column(df.columns)

    if not ndc_col:

        raise Exception("NDC column not found!")
 
    df[ndc_col] = df[ndc_col].astype(str).str.strip()

    df = df[df[ndc_col] != "nan"]

    df = df.set_index(ndc_col)
 
    return df
 
 
def normalize_text(text):

    if pd.isna(text):

        return ""

    return str(text).strip().lower().replace(" ", "").replace("-", "")
 
 
def find_ndc_column(columns):

    for col in columns:

        if "ndc" in col:

            return col

    return None
 
 
def try_float(val):

    if pd.isna(val):

        return None
 
    val = str(val).strip().replace("$", "").replace(",", "")
 
    if val in ("", "-"):

        return None
 
    try:

        return float(val)

    except:

        return None
 
 
def compare_data(df1, df2):

    mismatches = []

    group_summary = defaultdict(int)

    group_ndc_map = defaultdict(set)
 
    all_cols = sorted(set(df1.columns) & set(df2.columns))

    all_rows = sorted(set(df1.index) & set(df2.index))
 
    for col in all_cols:

        group = col.split("_")[0] if "_" in col else "Other"
 
        for row in all_rows:

            v1 = df1.at[row, col]

            v2 = df2.at[row, col]
 
            if pd.isna(v1) and pd.isna(v2):

                continue
 
            f1, f2 = try_float(v1), try_float(v2)
 
            if f1 is not None and f2 is not None:

                if abs(f1 - f2) > TOLERANCE:

                    mismatches.append((row, col, v1, v2))

                    group_summary[group] += 1

                    group_ndc_map[group].add(row)

            else:

                if str(v1).strip() != str(v2).strip():

                    mismatches.append((row, col, v1, v2))

                    group_summary[group] += 1

                    group_ndc_map[group].add(row)
 
    return mismatches, group_summary, group_ndc_map
 
 
def find_missing(df1, df2):

    missing_cols = sorted(set(df1.columns) - set(df2.columns))

    extra_cols = sorted(set(df2.columns) - set(df1.columns))

    missing_rows = sorted(set(df1.index) - set(df2.index))

    extra_rows = sorted(set(df2.index) - set(df1.index))
 
    return missing_cols, extra_cols, missing_rows, extra_rows
 
 
def generate_html(file1, file2, df1, df2, mismatches, group_summary,

                  group_ndc_map, missing_cols, extra_cols, missing_rows, extra_rows):
 
    file1_name = escape(os.path.basename(file1))

    file2_name = escape(os.path.basename(file2))
 
    html = []
 
    html.append("""
<html>
<head>
<style>

table { border-collapse: collapse; }

th, td { padding: 8px 12px; border: 1px solid black; text-align: left; }

th { background-color: #f2f2f2; }
</style>
</head>
<body>

""")
 
    html.append(f"<h1>Comparison Report</h1>")

    html.append(f"<h2>{file1_name} vs {file2_name}</h2>")

    html.append("<h3>Summary</h3>")

    html.append(f"<p>Rows: {len(df1)} vs {len(df2)}</p>")

    html.append(f"<p>Columns: {len(df1.columns)} vs {len(df2.columns)}</p>")

    html.append(f"<p>Total mismatches: {len(mismatches)}</p>")
 
    html.append("<h3>Mismatch by Group</h3>")

    html.append("<table><tr><th>Group</th><th>Mismatch Count</th><th>NDCs</th></tr>")
 
    for g in sorted(group_summary):

        ndcs = ", ".join(sorted(map(str, group_ndc_map[g])))

        html.append(f"<tr><td>{escape(g)}</td><td>{group_summary[g]}</td><td>{escape(ndcs)}</td></tr>")
 
    html.append("</table>")
 
    html.append("<h3>Detailed Mismatches</h3>")

    html.append("<table>")

    html.append(f"<tr><th>NDC</th><th>Column</th><th>{file1_name}</th><th>{file2_name}</th></tr>")
 
    for row, col, v1, v2 in mismatches:

        html.append(f"<tr><td>{row}</td><td>{escape(col)}</td><td>{v1}</td><td>{v2}</td></tr>")
 
    html.append("</table>")
 
    html.append("<h3>Missing / Extra Columns</h3>")

    html.append(f"<p>Missing Columns: {missing_cols}</p>")

    html.append(f"<p>Extra Columns: {extra_cols}</p>")
 
    html.append("<h3>Missing / Extra Rows (by NDC)</h3>")

    html.append(f"<p>Missing Rows: {missing_rows}</p>")

    html.append(f"<p>Extra Rows: {extra_rows}</p>")
 
    html.append("</body></html>")
 
    return "\n".join(html)
 
 
def main():
 
    app_path = Path(APP_FOLDER)

    excel_path = Path(EXCEL_FOLDER)

    out_path = Path(OUTPUT_FOLDER)

    out_path.mkdir(parents=True, exist_ok=True)
 
    # ===== LOAD MAPPING (same as RA Output) =====

    mapping = load_mapping()

    rpt_to_tc = reverse_mapping(mapping)
 
    app_map = {}

    excel_map = {}
 
    # ===== FILE1 (RPT BASED instead of TC*_App*) =====

    for f in app_path.glob("*.xlsx"):
 
        rpt_id = extract_rpt_id(f.name)
 
        if rpt_id and rpt_id in rpt_to_tc:
 
            tc = rpt_to_tc[rpt_id].replace("TC", "")

            app_map[tc] = f
 
    # ===== FILE2 (KEEP SAME AS WORKING LOGIC BUT FIX PATTERN) =====

    for f in excel_path.glob("*.xlsx"):
 
        tc = extract_tc(f.name)
 
        if tc:

            excel_map[tc] = f
 
    # ===== DEBUG (VERY IMPORTANT) =====

    print("\nChecking File1 mapping:")

    print(app_map)
 
    print("\nChecking File2 mapping:")

    print(excel_map)
 
    print("\nFile1 TC keys:", app_map.keys())

    print("File2 TC keys:", excel_map.keys())
 
    common_tcs = sorted(set(app_map.keys()) & set(excel_map.keys()), key=int)
 
    print(f"\nRunning {len(common_tcs)} scenarios...\n")
 
    for tc in common_tcs:
 
        APP_FILE = app_map[tc]

        EXCEL_FILE = excel_map[tc]
 
        OUTPUT_HTML = out_path / f"TC{tc}_Report.html"
 
        print(f"Processing TC{tc}...")
 
        df1 = load_and_prepare(APP_FILE)

        df2 = load_and_prepare(EXCEL_FILE)
 
        mismatches, group_summary, group_ndc_map = compare_data(df1, df2)
 
        missing_cols, extra_cols, missing_rows, extra_rows = find_missing(df1, df2)
 
        html = generate_html(

            APP_FILE, EXCEL_FILE, df1, df2,

            mismatches, group_summary, group_ndc_map,

            missing_cols, extra_cols, missing_rows, extra_rows

        )
 
        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:

            f.write(html)
 
        print(f"TC{tc} completed\n")
 
    print("All scenarios completed.")
 
 
 
if __name__ == "__main__":

    main()