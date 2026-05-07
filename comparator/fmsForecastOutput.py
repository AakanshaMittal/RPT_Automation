import pandas as pd

from html import escape

import re

import json

from pathlib import Path
 
BASE_DATA_PATH = "data"

APP_FOLDER = f"{BASE_DATA_PATH}/fmsForecastOutput/File1"

EXCEL_FOLDER = f"{BASE_DATA_PATH}/fmsForecastOutput/File2"

OUTPUT_FOLDER = f"{BASE_DATA_PATH}/fmsForecastOutput/Output"

MAPPING_FILE = "config/mapping.json"
 
def is_empty(x):

    return pd.isna(x) or str(x).strip() == ""
 
def clean(x):

    return "" if is_empty(x) else str(x).strip()
 
def normalize_key(x):

    return "" if not x else str(x).strip().lower()
 
def try_float(x):

    try:

        return float(str(x).replace(",", "").replace("%", ""))

    except:

        return None
 
def load_mapping():

    with open(MAPPING_FILE) as f:

        return json.load(f)
 
def reverse_mapping(mapping):

    return {v: k for k, v in mapping.items()}
 
def extract_rpt_id(filename):

    match = re.search(r"(RPT\d+[A-Z]{2})", filename.upper())

    return match.group(1) if match else None
 
def extract_tc(filename):

    match = re.search(r"TC(\d+)", filename.upper())

    return match.group(1) if match else None
 
def find_section_row(df):

    for i in range(15):

        row = " ".join(str(x).lower() for x in df.iloc[i] if not is_empty(x))

        if "no adjustment" in row:

            return i

    return 0
 
def find_group_row(df, section_idx):

    for i in range(section_idx + 1, section_idx + 6):

        row = " ".join(str(x).lower() for x in df.iloc[i] if not is_empty(x))

        if "rebate" in row or "script" in row:

            return i

    return section_idx + 1
 
def forward_fill_row(row):

    res, last = [], ""

    for v in row:

        if not is_empty(v):

            last = str(v).strip()

        res.append(last)

    return res
 
def get_column_blocks(df):

    blocks, start = [], None

    for c in range(df.shape[1]):

        if df.iloc[:, c].isna().all():

            if start is not None:

                blocks.append((start, c))

                start = None

        else:

            if start is None:

                start = c

    if start is not None:

        blocks.append((start, df.shape[1]))

    return blocks
 
def get_row_blocks(df):

    blocks, start = [], None

    for i in range(df.shape[0]):

        if all(is_empty(x) for x in df.iloc[i]):

            if start is not None:

                blocks.append((start, i))

                start = None

        else:

            if start is None:

                start = i

    if start is not None:

        blocks.append((start, df.shape[0]))

    return blocks
 
def extract_structure(df):

    structure = {}

    s_idx = find_section_row(df)

    g_idx = find_group_row(df, s_idx)

    section_row = forward_fill_row(df.iloc[s_idx])

    group_row = forward_fill_row(df.iloc[g_idx])

    col_blocks = get_column_blocks(df)
 
    last_section, last_group = "", ""
 
    for c_start, c_end in col_blocks:

        sec = normalize_key(section_row[c_start])

        grp = normalize_key(group_row[c_start])
 
        if sec:

            last_section = sec

        if grp:

            last_group = grp
 
        section = last_section or "unknown"

        group = last_group or "unknown"
 
        structure.setdefault(section, {}).setdefault(group, [])
 
        sub_df = df.iloc[:, c_start:c_end]
 
        for r_start, r_end in get_row_blocks(sub_df):

            table = sub_df.iloc[r_start:r_end].reset_index(drop=True)
 
            if table.dropna(how="all").empty:

                continue
 
            if table.shape[0] <= 1 and table.shape[1] <= 1:

                continue
 
            structure[section][group].append(table)
 
    return structure
 
def compare_tables(t1, t2, tol=0.0001):

    mismatches = []

    rows = max(len(t1), len(t2))

    cols = max(t1.shape[1], t2.shape[1])
 
    for i in range(rows):

        for j in range(cols):

            v1 = clean(t1.iloc[i, j] if i < len(t1) and j < t1.shape[1] else "")

            v2 = clean(t2.iloc[i, j] if i < len(t2) and j < t2.shape[1] else "")
 
            f1, f2 = try_float(v1), try_float(v2)
 
            if f1 is not None and f2 is not None:

                if abs(f1 - f2) > tol:

                    mismatches.append((i, j))

            else:

                if v1 != v2:

                    mismatches.append((i, j))
 
    return mismatches
 
def generate_html(results):

    html = "<html><body><h1>Mismatch Report</h1>"
 
    for sec, groups in results.items():

        html += f"<h2>{sec}</h2>"
 
        for grp, tables in groups.items():

            html += f"<h3>{grp}</h3>"
 
            for idx, (t1, t2, mm) in enumerate(tables):

                if not mm:

                    continue
 
                html += f"<h4>Table {idx+1}</h4><table border=1>"
 
                rows = max(len(t1), len(t2))

                cols = max(t1.shape[1], t2.shape[1])
 
                for i in range(rows):

                    html += "<tr>"

                    for j in range(cols):

                        v1 = clean(t1.iloc[i, j] if i < len(t1) else "")

                        v2 = clean(t2.iloc[i, j] if i < len(t2) else "")

                        cls = " style='background-color:#ffcccc'" if (i,j) in mm else ""

                        html += f"<td{cls}>{escape(v1)} | {escape(v2)}</td>"

                    html += "</tr>"
 
                html += "</table>"
 
    html += "</body></html>"

    return html
 
def main():

    app_path = Path(APP_FOLDER)

    excel_path = Path(EXCEL_FOLDER)

    out_path = Path(OUTPUT_FOLDER)

    out_path.mkdir(parents=True, exist_ok=True)
 
    mapping = load_mapping()

    rpt_to_tc = reverse_mapping(mapping)
 
    app_map = {}

    excel_map = {}
 
    for f in app_path.glob("*.xlsx"):

        rpt_id = extract_rpt_id(f.name)

        if rpt_id and rpt_id in rpt_to_tc:

            tc = rpt_to_tc[rpt_id].replace("TC", "")

            app_map[tc] = f
 
    for f in excel_path.glob("TC*_Exc_FMS*.xlsx"):

        tc = extract_tc(f.name)

        if tc:

            excel_map[tc] = f
 
    common_tcs = sorted(set(app_map.keys()) & set(excel_map.keys()), key=int)
 
    print(f"\nRunning {len(common_tcs)} scenarios...\n")
 
    for tc in common_tcs:

        df1 = pd.read_excel(app_map[tc], header=None)

        df2 = pd.read_excel(excel_map[tc], header=None)
 
        s1 = extract_structure(df1)

        s2 = extract_structure(df2)
 
        results = {}
 
        for sec in set(s1) | set(s2):

            results[sec] = {}

            for grp in set(s1.get(sec, {})) | set(s2.get(sec, {})):

                t1_list = s1.get(sec, {}).get(grp, [])

                t2_list = s2.get(sec, {}).get(grp, [])
 
                results[sec][grp] = []
 
                for i in range(max(len(t1_list), len(t2_list))):

                    t1 = t1_list[i] if i < len(t1_list) else pd.DataFrame()

                    t2 = t2_list[i] if i < len(t2_list) else pd.DataFrame()

                    mm = compare_tables(t1, t2)

                    results[sec][grp].append((t1, t2, mm))
 
        html = generate_html(results)
 
        output_file = out_path / f"TC{tc}_FMS_Report_Dev6May.html"
 
        with open(output_file, "w") as f:

            f.write(html)
 
        print(f"TC{tc} completed")
 
    print("All scenarios completed.")
 
if __name__ == "__main__":

    main()