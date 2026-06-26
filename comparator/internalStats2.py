import pandas as pd

from html import escape

import os

import json

import re

from pathlib import Path

from openpyxl.styles import PatternFill
 
BASE_PATH = Path("data/internalStats2")
 
FILE1_DIR = BASE_PATH / "File1"

FILE2_DIR = BASE_PATH / "RebateRefresh"

OUTPUT_DIR = BASE_PATH / "RebateRefreshOutput"

MAPPING_FILE = "config/mapping.json"
 
# Tolerance

TOLERANCE = 0.0001
 
os.makedirs(OUTPUT_DIR, exist_ok=True)
 
 
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

    return f"TC{match.group(1)}" if match else None
 
 
def is_empty(x):

    return pd.isna(x) or str(x).strip() == ""
 
 
def clean(x):

    return "" if is_empty(x) else str(x).strip()
 
 
def try_float(x):

    try:

        return float(str(x).replace(",", "").replace("%", ""))

    except:

        return None
 
 
def trim_table(df):

    df = df.dropna(how="all")

    df = df.loc[:, df.notna().any()]

    return df.reset_index(drop=True)
 
 
def is_numeric_column(col):

    count, total = 0, 0

    for val in col:

        if not is_empty(val):

            total += 1

            if try_float(val) is not None:

                count += 1

    return total > 0 and (count / total) > 0.6
 
 
def get_column_blocks(df):

    blocks, start = [], None

    c = 0
 
    while c < df.shape[1]:

        col = df.iloc[:, c]

        empty_ratio = col.isna().sum() / len(col)
 
        if empty_ratio > 0.95:

            left_numeric = is_numeric_column(df.iloc[:, c - 1]) if c - 1 >= 0 else False

            right_numeric = is_numeric_column(df.iloc[:, c + 1]) if c + 1 < df.shape[1] else False
 
            if left_numeric and right_numeric and start is not None:

                blocks.append((start, c))

                start = None
 
            c += 1

            continue
 
        if start is None:

            start = c
 
        c += 1
 
    if start is not None:

        blocks.append((start, df.shape[1]))
 
    return blocks
 
 
def get_row_blocks(df):

    blocks, start = [], None
 
    for i in range(df.shape[0]):

        if all(is_empty(x) for x in df.iloc[i]):

            if start is not None and i - start > 0:

                blocks.append((start, i))

            start = None

        else:

            if start is None:

                start = i
 
    if start is not None and df.shape[0] - start > 0:

        blocks.append((start, df.shape[0]))
 
    return blocks
 
 
def extract_tables(df):

    tables = {}
 
    for c_start, c_end in get_column_blocks(df):

        sub_df = df.iloc[:, c_start:c_end]
 
        for r_start, r_end in get_row_blocks(sub_df):

            table = trim_table(sub_df.iloc[r_start:r_end].copy())
 
            if table.shape[0] < 2:

                continue
 
            name = ""

            for val in table.iloc[0]:

                val_clean = clean(val)

                if val_clean:

                    name = val_clean

                    break
 
            if not name:

                continue
 
            key = name.lower()
 
            if key in tables:

                key = f"{key}_{len(tables)}"
 
            tables[key] = table.reset_index(drop=True)
 
    return tables
 
 
def get_drug_name_column(df):

    for col in range(df.shape[1]):

        for row in range(min(3, df.shape[0])):

            if "drug name" in str(df.iloc[row, col]).lower():

                return col

    return None
 
 
def align_tables(t1, t2):

    k1, k2 = get_drug_name_column(t1), get_drug_name_column(t2)
 
    if k1 is None or k2 is None:

        return t1, t2
 
    def extract_data(df, c):

        data, header = {}, 0
 
        for i in range(min(3, len(df))):

            if "drug name" in str(df.iloc[i, c]).lower():

                header = i

                break
 
        for i in range(header + 1, len(df)):

            drug = clean(df.iloc[i, c])

            if drug:

                data[drug] = [clean(df.iloc[i, j]) for j in range(df.shape[1])]
 
        return data
 
    d1, d2 = extract_data(t1, k1), extract_data(t2, k2)
 
    all_drugs = sorted(set(d1) | set(d2))

    max_cols = max(t1.shape[1], t2.shape[1])
 
    a1, a2 = [], []
 
    for drug in all_drugs:

        r1 = d1.get(drug, [""] * max_cols)

        r2 = d2.get(drug, [""] * max_cols)
 
        a1.append(r1 + [""] * (max_cols - len(r1)))

        a2.append(r2 + [""] * (max_cols - len(r2)))
 
    return pd.DataFrame(a1), pd.DataFrame(a2)
 
 
def normalize_text(v):

    """

    Rule:

    If both values contain MFP anywhere:

      MFP vs MFP 2026 / MFP 2027 / etc = PASS

    """

    txt = clean(v).upper()
 
    if txt.startswith("MFP"):

        return "MFP"
 
    return txt
 
 
def compare_tables(t1, t2, tol=TOLERANCE):

    mismatches = []
 
    rows = min(len(t1), len(t2))

    cols = min(t1.shape[1], t2.shape[1])
 
    for i in range(rows):

        for j in range(cols):

            v1 = clean(t1.iloc[i, j])

            v2 = clean(t2.iloc[i, j])
 
            f1 = try_float(v1)

            f2 = try_float(v2)
 
            # Numeric compare with tolerance

            if f1 is not None and f2 is not None:

                if abs(f1 - f2) > tol:

                    mismatches.append((i, j))

                continue
 
            # MFP handling

            n1 = normalize_text(v1)

            n2 = normalize_text(v2)
 
            if n1 != n2:

                mismatches.append((i, j))
 
    return mismatches
 
 
def generate_html(results, f1_name, f2_name, output_path):

    html = """
<html>
<head>
<style>

body {font-family: Arial;}

table {border-collapse: collapse; margin-bottom:20px;}

td, th {border:1px solid black; padding:6px;}

.mismatch {background:#ffcccc;}

.flex {display:flex; gap:40px;}
</style>
</head>
<body>
 
<h1>Detailed Stats Internal 2</h1>
 
<table>
<tr>
<th>Table</th>
<th>Mismatches</th>
</tr>

"""
 
    for name, (_, _, mm) in results.items():

        html += f"<tr><td>{name}</td><td>{len(mm)}</td></tr>"
 
    html += "</table>"
 
    for name, (t1, t2, mm) in results.items():

        if not mm:

            continue
 
        html += f"<h3>{name}</h3><div class='flex'>"
 
        for df, label in [(t1, f1_name), (t2, f2_name)]:

            html += f"<div><h4>{label}</h4><table>"
 
            for i in range(len(df)):

                html += "<tr>"
 
                for j in range(df.shape[1]):

                    cls = "mismatch" if (i, j) in mm else ""

                    html += f"<td class='{cls}'>{escape(clean(df.iloc[i, j]))}</td>"
 
                html += "</tr>"
 
            html += "</table></div>"
 
        html += "</div>"
 
    html += "</body></html>"
 
    with open(output_path, "w", encoding="utf-8") as f:

        f.write(html)
 
 
def generate_consolidated_excel(summary, all_results, output_path):

    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
 
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
 
        summary_df = pd.DataFrame(summary).fillna(0).astype(int)

        summary_df.index.name = "Table"

        summary_df.reset_index(inplace=True)
 
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
 
        for tc, results in all_results.items():

            sheet = tc[:31]

            row = 0
 
            ws = writer.book.create_sheet(sheet)

            writer.sheets[sheet] = ws
 
            for name, (t1, t2, mm) in results.items():

                if not mm:

                    continue
 
                ws.cell(row=row + 1, column=1, value=name)
 
                t1.to_excel(

                    writer,

                    sheet_name=sheet,

                    startrow=row + 2,

                    startcol=0,

                    index=False,

                    header=False

                )
 
                offset = t1.shape[1] + 3
 
                t2.to_excel(

                    writer,

                    sheet_name=sheet,

                    startrow=row + 2,

                    startcol=offset,

                    index=False,

                    header=False

                )
 
                for i, j in mm:

                    ws.cell(row=row + 3 + i, column=1 + j).fill = red_fill

                    ws.cell(row=row + 3 + i, column=offset + 1 + j).fill = red_fill
 
                row += max(len(t1), 3) + 5
 
        if "Sheet" in writer.book.sheetnames:

            del writer.book["Sheet"]
 
 
def main():

    mapping = load_mapping()

    rev_map = reverse_mapping(mapping)
 
    file1_dict = {}

    file2_dict = {}
 
    for f in FILE1_DIR.glob("*.xlsx"):

        rpt = extract_rpt_id(f.name)
 
        if rpt and rpt in rev_map:

            file1_dict[rev_map[rpt]] = f
 
    for f in FILE2_DIR.glob("*.xlsx"):

        tc = extract_tc(f.name)
 
        if tc:

            file2_dict[tc] = f
 
    summary = {}

    all_results = {}
 
    common_tcs = sorted(set(file1_dict) & set(file2_dict))
 
    for tc in common_tcs:

        print(f"Processing {tc}...")
 
        df1 = pd.read_excel(file1_dict[tc], header=None)

        df2 = pd.read_excel(file2_dict[tc], header=None)
 
        t1_map = extract_tables(df1)

        t2_map = extract_tables(df2)
 
        results = {}

        summary[tc] = {}
 
        for key in set(t1_map) & set(t2_map):

            t1, t2 = align_tables(t1_map[key], t2_map[key])
 
            mm = compare_tables(t1, t2, TOLERANCE)
 
            results[key] = (t1, t2, mm)

            summary[tc][key] = len(mm)
 
        if results:

            all_results[tc] = results
 
            generate_html(

                results,

                file1_dict[tc].name,

                file2_dict[tc].name,

                OUTPUT_DIR / f"{tc}_Internal2_26Jun.html"

            )
 
    generate_consolidated_excel(

        summary,

        all_results,

        OUTPUT_DIR / "Consolidated_Report_26Jun.xlsx"

    )
 
    print("All scenarios completed.")
 
 
if __name__ == "__main__":

    main()
 