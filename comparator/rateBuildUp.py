import pandas as pd

import re

import json

from pathlib import Path

from collections import defaultdict
 
BASE_DATA_PATH = "data"
 
FILE1_FOLDER = f"{BASE_DATA_PATH}/rateBuildUp/File1"

FILE2_FOLDER = f"{BASE_DATA_PATH}/rateBuildUp/File2"

OUTPUT_FOLDER = f"{BASE_DATA_PATH}/rateBuildUp/Output"
 
MAPPING_FILE = "config/mapping.json"
 
TOLERANCE = 0.0001
 
# =========================

# BASIC HELPERS

# =========================
 
def is_blank(x):
 
    return pd.isna(x) or str(x).strip() == ""
 
 
def is_number(x):
 
    try:

        float(str(x).replace(",", ""))

        return True

    except:

        return False
 
 
def to_float(x):
 
    try:

        return float(str(x).replace(",", ""))

    except:

        return None
 
 
# =========================

# MISMATCH CLASSIFICATION

# =========================
 
def classify_difference(diff):
 
    if diff >= 1:

        return "Integer"
 
    elif diff >= 0.1:

        return "1st Decimal"
 
    elif diff >= 0.01:

        return "2nd Decimal"
 
    elif diff >= 0.001:

        return "3rd Decimal"
 
    elif diff >= TOLERANCE:

        return "4th Decimal"
 
    return "OK"
 
 
def calculate_status(details):
 
    levels = [d["category"] for d in details]
 
    if "Integer" in levels or "1st Decimal" in levels:

        return "FAIL"
 
    elif levels:

        return "REVIEW"
 
    return "PASS"
 
 
# =========================

# COLUMN LIMIT (CG)

# =========================
 
def excel_col_to_index(col):
 
    col = col.upper()
 
    result = 0
 
    for c in col:

        result = result * 26 + (ord(c) - ord('A') + 1)
 
    return result - 1
 
 
MAX_VALID_COL = excel_col_to_index("CG")
 
# =========================

# PERIOD LOGIC

# =========================
 
def is_period_value(x):
 
    s = str(x).strip()
 
    if not s:

        return False
 
    if re.match(r"\d{2}/\d{2}\s*-\s*\d{2}/\d{2}", s):

        return True
 
    if s.isdigit():

        return 2000 <= int(s) <= 2100
 
    return False
 
 
def is_period_row(row):
 
    return sum(

        1 for i, v in enumerate(row)

        if i <= MAX_VALID_COL and is_period_value(v)

    ) >= 2
 
 
# =========================

# FILE HELPERS

# =========================
 
def read_xlsx(path):
 
    df = pd.read_excel(path, header=None)
 
    return df.fillna("").values.tolist()
 
 
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
 
 
# =========================

# LEFT SIDE LOGIC

# =========================
 
def detect_sections(grid, section_row):
 
    section_by_col = {}
 
    last = ""
 
    rows = len(grid)
 
    for c in range(MAX_VALID_COL + 1):
 
        col_has_any_value = any(

            not is_blank(grid[r][c]) for r in range(rows)

        )
 
        if not col_has_any_value:
 
            last = ""
 
            section_by_col[c] = ""
 
            continue
 
        val = grid[section_row][c]
 
        if not is_blank(val):

            last = str(val).strip()
 
        section_by_col[c] = last
 
    return section_by_col
 
 
def find_tables(grid, start_row):
 
    tables = []
 
    r = start_row
 
    cols = range(len(grid[0]))
 
    while r < len(grid):
 
        if all(is_blank(grid[r][c]) for c in cols):
 
            r += 1
 
            continue
 
        top = r
 
        while r < len(grid) and not all(

            is_blank(grid[r][c]) for c in cols

        ):

            r += 1
 
        tables.append((top, r - 1))
 
        r += 1
 
    return tables
 
 
def forward_fill_row(row):
 
    out, last = [], ""
 
    for v in row:
 
        if not is_blank(v):

            last = str(v).strip()
 
        out.append(last)
 
    return out
 
 
def extract_records(grid, source, section_row):
 
    section_map = detect_sections(grid, section_row)
 
    records = []
 
    tables = find_tables(grid, section_row + 1)
 
    for top, bottom in tables:
 
        header_rows = []
 
        r = top
 
        while r <= bottom and len(header_rows) < 3:
 
            if not all(is_blank(grid[r][c]) for c in range(len(grid[0]))):

                header_rows.append(r)
 
            r += 1
 
        if len(header_rows) < 3:

            continue
 
        g_row, sg_row, p_row = header_rows[:3]
 
        if not is_period_row(grid[p_row]):

            continue
 
        group = forward_fill_row(grid[g_row])
 
        sub_group = forward_fill_row(grid[sg_row])
 
        period = forward_fill_row(grid[p_row])
 
        for r in range(p_row + 1, bottom + 1):
 
            row = grid[r]
 
            for c in range(MAX_VALID_COL + 1):
 
                if is_number(row[c]):
 
                    records.append({
 
                        "source": source,
 
                        "section": section_map.get(c, ""),
 
                        "group": group[c],
 
                        "sub_group": sub_group[c],
 
                        "period": period[c],
 
                        "channel": f"Row{r}",
 
                        "value": to_float(row[c])
 
                    })
 
    return records
 
 
# =========================

# COMPARE (LEFT)

# =========================
 
def compare(records):
 
    data = defaultdict(dict)
 
    mismatches = set()
 
    mismatch_details = []
 
    for r in records:
 
        key = (

            r["section"],

            r["group"],

            r["sub_group"],

            r["period"],

            r["channel"]

        )
 
        data[key][r["source"]] = r["value"]
 
    for k, v in data.items():
 
        if len(v) == 2:
 
            a, b = list(v.values())
 
            if a is not None and b is not None:
 
                diff = abs(a - b)
 
                if diff > TOLERANCE:
 
                    mismatches.add(k)
 
                    mismatch_details.append({
 
                        "key": k,
 
                        "diff": diff,
 
                        "category": classify_difference(diff)
 
                    })
 
    return data, mismatches, mismatch_details
 
 
def has_issue(sec, grp, sg, data, mismatches, sources):
 
    for (s, g, subg, p, ch), vals in data.items():
 
        if s == sec and g == grp and subg == sg:
 
            if (s, g, subg, p, ch) in mismatches:

                return True
 
            if len(vals) < len(sources):

                return True
 
    return False
 
 
# =========================

# RIGHT SIDE TABLES

# =========================
 
def find_right_side_tables(grid):
 
    tables = []
 
    rows, cols = len(grid), len(grid[0])
 
    visited = set()
 
    for r in range(rows):
 
        for c in range(MAX_VALID_COL + 1, cols):
 
            if (r, c) in visited or is_blank(grid[r][c]):

                continue
 
            stack = [(r, c)]
 
            min_r = max_r = r
 
            min_c = max_c = c
 
            while stack:
 
                cr, cc = stack.pop()
 
                if (cr, cc) in visited:

                    continue
 
                if cr < 0 or cr >= rows or cc < 0 or cc >= cols:

                    continue
 
                if is_blank(grid[cr][cc]):

                    continue
 
                visited.add((cr, cc))
 
                min_r, max_r = min(min_r, cr), max(max_r, cr)
 
                min_c, max_c = min(min_c, cc), max(max_c, cc)
 
                for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:

                    stack.append((cr+dr, cc+dc))
 
            tables.append((min_r, max_r, min_c, max_c))
 
    return tables
 
 
def extract_positional_table(grid, bounds):
 
    min_r, max_r, min_c, max_c = bounds
 
    table = []
 
    for r in range(min_r, max_r + 1):
 
        row = []
 
        for c in range(min_c, max_c + 1):
 
            val = grid[r][c]
 
            row.append(to_float(val) if is_number(val) else val)
 
        table.append(row)
 
    return table
 
 
def compare_positional_tables(g1, g2, tables):
 
    results = []
 
    right_side_details = []
 
    for bounds in tables:
 
        t1 = extract_positional_table(g1, bounds)
 
        t2 = extract_positional_table(g2, bounds)
 
        mismatches = []
 
        for r in range(len(t1)):
 
            for c in range(len(t1[0])):
 
                v1, v2 = t1[r][c], t2[r][c]
 
                if isinstance(v1, float) and isinstance(v2, float):
 
                    diff = abs(v1 - v2)
 
                    if diff > TOLERANCE:
 
                        mismatches.append((r, c))
 
                        right_side_details.append({
 
                            "table": bounds,
 
                            "row": r,
 
                            "col": c,
 
                            "diff": diff,
 
                            "category": classify_difference(diff)
 
                        })
 
                elif v1 != v2:
 
                    mismatches.append((r, c))
 
        if mismatches:
 
            results.append((bounds, t1, t2, mismatches))
 
    return results, right_side_details
 
 
# =========================

# HTML

# =========================
 
def generate_html(

    data,

    mismatches,

    sources,

    pos_results,

    all_details,

    status

):
 
    html = []
 
    html.append("<html><style>")
 
    html.append("""
 
    body{font-family:Arial}
 
    table{border-collapse:collapse;margin:10px}
 
    th,td{border:1px solid #aaa;padding:6px}
 
    .integer{background:#ffe0e0}
 
    .first{background:#ffe0e0}
 
    .second{background:#ffe699}
 
    .third{background:#fff2cc}
 
    .fourth{background:#fff2cc}
 
    .missing{background:#e0e0e0;color:#888}
 
    .wrap{display:flex;gap:20px}
 
    """)
 
    html.append("</style><body>")
 
    counts = defaultdict(int)
 
    for d in all_details:

        counts[d["category"]] += 1
 
    html.append("<h1>Execution Summary</h1>")
 
    html.append("<table>")
 
    html.append("""
 
    <tr>
<th>Status</th>
<th>Integer</th>
<th>1st Decimal</th>
<th>2nd Decimal</th>
<th>3rd Decimal</th>
<th>4th Decimal</th>
</tr>
 
    """)
 
    html.append(f"""
 
    <tr>
<td>{status}</td>
<td>{counts['Integer']}</td>
<td>{counts['1st Decimal']}</td>
<td>{counts['2nd Decimal']}</td>
<td>{counts['3rd Decimal']}</td>
<td>{counts['4th Decimal']}</td>
</tr>
 
    """)
 
    html.append("</table>")
 
    summary = defaultdict(int)
 
    for (s, g, sg, p, ch) in mismatches:

        summary[(s, g, sg)] += 1
 
    right_side_count = len(pos_results)
 
    html.append("<h2>Mismatch Summary</h2>")
 
    html.append("<table>")
 
    html.append("""
 
    <tr>
<th>Section</th>
<th>Group</th>
<th>Sub-Group</th>
<th>Count</th>
</tr>
 
    """)
 
    for (sec, grp, sg), count in summary.items():
 
        html.append(f"""
 
        <tr>
<td>{sec}</td>
<td>{grp}</td>
<td>{sg}</td>
<td>{count}</td>
</tr>
 
        """)
 
    html.append(f"""
 
    <tr>
<td colspan='3'><b>Right Side Tables</b></td>
<td><b>{right_side_count}</b></td>
</tr>
 
    """)
 
    html.append("</table>")
 
    structure = defaultdict(

        lambda: defaultdict(

            lambda: defaultdict(

                lambda: defaultdict(dict)

            )

        )

    )
 
    periods = defaultdict(set)
 
    mismatch_map = {}
 
    for d in all_details:
 
        if "key" in d:

            mismatch_map[d["key"]] = d["category"]
 
    for (sec, grp, sg, p, ch), vals in data.items():
 
        periods[(sec, grp, sg)].add(p)
 
        for src, val in vals.items():
 
            structure[sec][grp][sg][src].setdefault(ch, {})[p] = val
 
    for sec in structure:
 
        for grp in structure[sec]:
 
            for sg in structure[sec][grp]:
 
                if not has_issue(

                    sec,

                    grp,

                    sg,

                    data,

                    mismatches,

                    sources

                ):

                    continue
 
                html.append(f"<h2>{sec}</h2>")

                html.append(f"<h3>{grp}</h3>")

                html.append(f"<h4>{sg}</h4>")
 
                ps = sorted(periods[(sec, grp, sg)])
 
                html.append("<div class='wrap'>")
 
                for src in sources:
 
                    html.append("<table>")
 
                    html.append(

                        f"<tr><th colspan='{len(ps)+1}'>{src}</th></tr>"

                    )
 
                    html.append(

                        "<tr><th>Channel</th>"

                        + "".join(f"<th>{p}</th>" for p in ps)

                        + "</tr>"

                    )
 
                    for ch in structure[sec][grp][sg][src]:
 
                        html.append(f"<tr><td>{ch}</td>")
 
                        for p in ps:
 
                            val = structure[sec][grp][sg][src][ch].get(p)
 
                            key = (sec, grp, sg, p, ch)
 
                            if val is None:
 
                                html.append(

                                    "<td class='missing'>-</td>"

                                )
 
                            elif key in mismatch_map:
 
                                cat = mismatch_map[key]
 
                                cls = {

                                    "Integer":"integer",

                                    "1st Decimal":"first",

                                    "2nd Decimal":"second",

                                    "3rd Decimal":"third",

                                    "4th Decimal":"fourth"

                                }[cat]
 
                                html.append(

                                    f"<td class='{cls}'>{val}</td>"

                                )
 
                            else:
 
                                html.append(f"<td>{val}</td>")
 
                        html.append("</tr>")
 
                    html.append("</table>")
 
                html.append("</div>")
 
    if pos_results:
 
        html.append("<h2>Independent Tables (Right Side)</h2>")
 
        for bounds, t1, t2, mism in pos_results:
 
            html.append("<div class='wrap'>")
 
            for table in [t1, t2]:
 
                html.append("<table>")
 
                for r in range(len(table)):
 
                    html.append("<tr>")
 
                    for c in range(len(table[0])):
 
                        val = table[r][c]
 
                        if (r, c) in mism:
 
                            html.append(

                                f"<td class='first'>{val}</td>"

                            )
 
                        else:
 
                            html.append(f"<td>{val}</td>")
 
                    html.append("</tr>")
 
                html.append("</table>")
 
            html.append("</div>")
 
    html.append("</body></html>")
 
    return "".join(html)
 
 
# =========================

# MAIN

# =========================
 
def main():
 
    section_row = 5
 
    f1_path = Path(FILE1_FOLDER)
 
    f2_path = Path(FILE2_FOLDER)
 
    out_path = Path(OUTPUT_FOLDER)
 
    out_path.mkdir(parents=True, exist_ok=True)
 
    mapping = load_mapping()
 
    rpt_to_tc = reverse_mapping(mapping)
 
    file1_map, file2_map = {}, {}
 
    summary_rows = []
 
    for f in f1_path.glob("*.xlsx"):
 
        rpt_id = extract_rpt_id(f.name)
 
        if rpt_id and rpt_id in rpt_to_tc:
 
            tc = rpt_to_tc[rpt_id].replace("TC", "")
 
            file1_map[tc] = f
 
    for f in f2_path.glob("*.xlsx"):
 
        tc = extract_tc(f.name)
 
        if tc:

            file2_map[tc] = f
 
    common_tcs = sorted(set(file1_map) & set(file2_map), key=int)
 
    for tc in common_tcs:
 
        file1, file2 = file1_map[tc], file2_map[tc]
 
        g1 = read_xlsx(file1)
 
        g2 = read_xlsx(file2)
 
        r1 = extract_records(g1, file1.name, section_row)
 
        r2 = extract_records(g2, file2.name, section_row)
 
        data, mismatches, mismatch_details = compare(r1 + r2)
 
        right_tables = find_right_side_tables(g1)
 
        pos_results, right_side_details = compare_positional_tables(

            g1,

            g2,

            right_tables

        )
 
        all_details = mismatch_details + right_side_details
 
        status = calculate_status(all_details)
 
        counts = defaultdict(int)
 
        for d in all_details:

            counts[d["category"]] += 1
 
        summary_rows.append({
 
            "TC": tc,
 
            "Status": status,
 
            "Integer": counts["Integer"],
 
            "1st Decimal": counts["1st Decimal"],
 
            "2nd Decimal": counts["2nd Decimal"],
 
            "3rd Decimal": counts["3rd Decimal"],
 
            "4th Decimal": counts["4th Decimal"]
 
        })
 
        html = generate_html(

            data,

            mismatches,

            [file1.name, file2.name],

            pos_results,

            all_details,

            status

        )
 
        with open(

            out_path / f"TC{tc}_RBU_Report_Stg15Jun.html",

            "w",

            encoding="utf-8"

        ) as f:
 
            f.write(html)
 
        print(f"TC{tc} done")
 
    summary_df = pd.DataFrame(summary_rows)
 
    summary_df.to_excel(

        out_path / "RBU_Execution_SummaryStg15Jun.xlsx",

        index=False

    )
 
    print("Execution Summary Excel Generated")
 
 
if __name__ == "__main__":
 
    main()
 