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

    return sum(1 for i, v in enumerate(row) if i <= MAX_VALID_COL and is_period_value(v)) >= 2
 
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

        col_has_any_value = any(not is_blank(grid[r][c]) for r in range(rows))

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

        while r < len(grid) and not all(is_blank(grid[r][c]) for c in cols):

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
 
    for r in records:

        key = (r["section"], r["group"], r["sub_group"], r["period"], r["channel"])

        data[key][r["source"]] = r["value"]
 
    for k, v in data.items():

        if len(v) == 2:

            a, b = list(v.values())

            if a is not None and b is not None:

                if abs(a - b) > TOLERANCE:

                    mismatches.add(k)
 
    return data, mismatches
 
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

    for bounds in tables:

        t1 = extract_positional_table(g1, bounds)

        t2 = extract_positional_table(g2, bounds)
 
        mismatches = []

        for r in range(len(t1)):

            for c in range(len(t1[0])):

                v1, v2 = t1[r][c], t2[r][c]
 
                if isinstance(v1, float) and isinstance(v2, float):

                    if abs(v1 - v2) > TOLERANCE:

                        mismatches.append((r, c))

                elif v1 != v2:

                    mismatches.append((r, c))
 
        if mismatches:

            results.append((bounds, t1, t2, mismatches))
 
    return results
 
# =========================

# HTML (FINAL MERGED)

# =========================

def generate_html(data, mismatches, sources, pos_results):

    html = []

    html.append("<html><style>")

    html.append("""

    body{font-family:Arial}

    table{border-collapse:collapse;margin:10px}

    th,td{border:1px solid #aaa;padding:6px}

    .bad{background:#ffe0e0}

    .missing{background:#f0f0f0;color:#999}

    .wrap{display:flex;gap:20px}

    """)

    html.append("</style><body>")
 
    # SUMMARY

    summary = defaultdict(int)

    for (s, g, sg, p, ch) in mismatches:

        summary[(s, g, sg)] += 1
 
    html.append("<h2>Mismatch Summary</h2><table>")

    html.append("<tr><th>Section</th><th>Group</th><th>Sub-Group</th><th>Count</th></tr>")
 
    for (sec, grp, sg), count in summary.items():

        html.append(f"<tr><td>{sec}</td><td>{grp}</td><td>{sg}</td><td>{count}</td></tr>")
 
    html.append("</table>")
 
    # LEFT TABLES

    structure = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(dict))))

    periods = defaultdict(set)
 
    for (sec, grp, sg, p, ch), vals in data.items():

        periods[(sec, grp, sg)].add(p)

        for src, val in vals.items():

            structure[sec][grp][sg][src].setdefault(ch, {})[p] = val
 
    for sec in structure:

        for grp in structure[sec]:

            for sg in structure[sec][grp]:

                if not has_issue(sec, grp, sg, data, mismatches, sources):

                    continue
 
                html.append(f"<h2>{sec}</h2><h3>{grp}</h3><h4>{sg}</h4>")

                ps = sorted(periods[(sec, grp, sg)])

                html.append("<div class='wrap'>")
 
                for src in sources:

                    html.append("<table>")

                    html.append(f"<tr><th colspan='{len(ps)+1}'>{src}</th></tr>")

                    html.append("<tr><th>Channel</th>" + "".join(f"<th>{p}</th>" for p in ps) + "</tr>")
 
                    for ch in structure[sec][grp][sg][src]:

                        html.append("<tr><td>{}</td>".format(ch))

                        for p in ps:

                            val = structure[sec][grp][sg][src][ch].get(p)

                            key = (sec, grp, sg, p, ch)
 
                            if val is None:

                                html.append("<td class='missing'>-</td>")

                            elif key in mismatches:

                                html.append(f"<td class='bad'>{val}</td>")

                            else:

                                html.append(f"<td>{val}</td>")
 
                        html.append("</tr>")

                    html.append("</table>")
 
                html.append("</div>")
 
    # RIGHT TABLES

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

                            html.append(f"<td class='bad'>{val}</td>")

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
 
        data, mismatches = compare(r1 + r2)
 
        right_tables = find_right_side_tables(g1)

        pos_results = compare_positional_tables(g1, g2, right_tables)
 
        html = generate_html(data, mismatches, [file1.name, file2.name], pos_results)
 
        with open(out_path / f"TC{tc}_RBU_Report_Dev24Apr.html", "w", encoding="utf-8") as f:

            f.write(html)
 
        print(f"TC{tc} done")
 
if __name__ == "__main__":

    main()
 