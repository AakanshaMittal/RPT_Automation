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
 
def detect_sections(grid, section_row):

    section_by_col = {}

    last = ""

    rows = len(grid)

    cols = len(grid[0])
 
    for c in range(cols):

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

    out = []

    last = ""
 
    for v in row:

        if not is_blank(v):

            last = str(v).strip()

        out.append(last)
 
    return out
 
def get_all_channels_from_table(grid, top, bottom):

    channels = []

    data_start = None
 
    for r in range(top, bottom + 1):

        if any(is_number(grid[r][c]) for c in range(len(grid[0]))):

            data_start = r

            break
 
    if data_start is None:

        return channels
 
    for r in range(data_start, bottom + 1):

        row = grid[r]
 
        for c in range(len(row)):

            if is_number(row[c]):

                channel_col = c - 1
 
                if channel_col >= 0:

                    channel_val = str(row[channel_col]).strip()
 
                    if channel_val == "":

                        row_has_total = any("total" in str(cell).lower() for cell in row)
 
                        if row_has_total or r == bottom:

                            channel_val = "Total"

                        else:

                            next_row_has_channel = False
 
                            if r + 1 <= bottom:

                                for nc in range(channel_col, len(grid[r+1])):

                                    if not is_blank(grid[r+1][nc]) and not is_number(grid[r+1][nc]):

                                        next_row_has_channel = True

                                        break
 
                            if not next_row_has_channel and len(channels) > 0:

                                channel_val = "Total"

                            else:

                                channel_val = "Unknown"
 
                    channels.append(channel_val)

                break
 
    return channels
 
def extract_records(grid, source, section_row):

    section_map = detect_sections(grid, section_row)

    records = []

    tables = find_tables(grid, section_row + 1)
 
    for top, bottom in tables:

        header_rows = []

        data_start = None
 
        for r in range(top, bottom + 1):

            if any(is_number(grid[r][c]) for c in range(len(grid[0]))):

                data_start = r

                break

            header_rows.append(r)
 
        if len(header_rows) < 3:

            continue
 
        g_row, sg_row, p_row = header_rows[:3]
 
        group = forward_fill_row(grid[g_row])

        sub_group = forward_fill_row(grid[sg_row])

        period = forward_fill_row(grid[p_row])
 
        all_channels = get_all_channels_from_table(grid, top, bottom)
 
        channel_by_row = {}

        row_idx = data_start
 
        for ch in all_channels:

            if row_idx <= bottom:

                channel_by_row[row_idx] = ch

                row_idx += 1
 
        for r in range(data_start, bottom + 1):

            row = grid[r]

            channel = channel_by_row.get(r, "")
 
            if channel == "" and r == bottom:

                channel = "Total"

            elif channel == "":

                if r > data_start and channel_by_row.get(r-1, "") == "Total":

                    channel = "Total"

                else:

                    continue
 
            num_cols = [c for c in range(len(row)) if is_number(row[c])]
 
            if not num_cols:

                for c in range(len(period)):

                    if period[c] and section_map.get(c):

                        if not is_blank(period[c]):

                            records.append({

                                "source": source,

                                "section": section_map[c],

                                "group": group[c] if c < len(group) else "",

                                "sub_group": sub_group[c] if c < len(sub_group) else "",

                                "period": period[c] if c < len(period) else "",

                                "channel": channel,

                                "value": None

                            })

                continue
 
            for c in num_cols:

                records.append({

                    "source": source,

                    "section": section_map[c],

                    "group": group[c] if c < len(group) else "",

                    "sub_group": sub_group[c] if c < len(sub_group) else "",

                    "period": period[c] if c < len(period) else "",

                    "channel": channel,

                    "value": to_float(row[c])

                })
 
    return records
 
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

                if abs(a - b) > 0.00001:

                    mismatches.add(k)
 
    return data, mismatches

def has_issue(sec, grp, sg, data, mismatches, sources):

    for (s, g, subg, p, ch), vals in data.items():

        if s == sec and g == grp and subg == sg:

            if (s, g, subg, p, ch) in mismatches:

                return True

            # Check if any channel has less than all sources

            # This ensures tables with partial data are shown

            if len(vals) < len(sources):

                return True

    return False
 
def generate_html(data, mismatches, sources):

    html = []

    html.append("<html><style>")

    html.append("""

    body{font-family:Arial}

    table{border-collapse:collapse;margin:10px}

    th,td{border:1px solid #aaa;padding:6px}

    .bad{background:#ffe0e0}

    .missing{background:#f0f0f0; color:#999}

    .wrap{display:flex;gap:20px}

    """)

    html.append("</style><body>")

    # SUMMARY TABLE

    summary = defaultdict(int)

    for (s, g, sg, p, ch) in mismatches:

        summary[(s, g, sg)] += 1

    html.append("<h2>Mismatch Summary</h2>")

    html.append("<table>")

    html.append("<tr><th>Section</th><th>Group</th><th>Sub-Group</th><th>Mismatch Count</th></tr>")

    for sec in sorted(set(k[0] for k in summary)):

        for grp in sorted(set(k[1] for k in summary if k[0] == sec)):

            for sg in sorted(set(k[2] for k in summary if k[0] == sec and k[1] == grp)):

                count = summary[(sec, grp, sg)]

                html.append(f"<tr><td>{sec}</td><td>{grp}</td><td>{sg}</td><td>{count}</td></tr>")

    html.append("</table>")

    # STRUCTURE

    structure = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(dict))))

    periods = defaultdict(set)

    channels_by_table = defaultdict(set)

    for (sec, grp, sg, p, ch), vals in data.items():

        periods[(sec, grp, sg)].add(p)

        channels_by_table[(sec, grp, sg)].add(ch)

        for src, val in vals.items():

            structure[sec][grp][sg][src].setdefault(ch, {})[p] = val

    # FILTERED TABLES

    for sec in sorted(structure.keys()):

        for grp in sorted(structure[sec].keys()):

            for sg in sorted(structure[sec][grp].keys()):

                if not has_issue(sec, grp, sg, data, mismatches, sources):

                    continue

                html.append(f"<h2>Section: {sec}</h2>")

                html.append(f"<h3>Group: {grp}</h3>")

                html.append(f"<h4>Sub-Group: {sg}</h4>")

                ps = sorted(periods[(sec, grp, sg)])

                html.append("<div class='wrap'>")

                # Get all channels that appear in either source

                all_channels = set()

                for src in sources:

                    if src in structure[sec][grp][sg]:

                        all_channels.update(structure[sec][grp][sg][src].keys())

                # Add channels that are in the table definition but might not have data

                all_channels.update(channels_by_table[(sec, grp, sg)])

                # Sort channels with Total at the end

                sorted_channels = sorted(all_channels, key=lambda x: (x != "Total", x))

                for src in sources:

                    html.append("<table>")

                    html.append(f"<tr><th colspan='{len(ps)+1}'>{src}</th></tr>")

                    html.append("<tr><th>Channel</th>" + "".join(f"<th>{p}</th>" for p in ps) + "</tr>")

                    for ch in sorted_channels:

                        html.append("<tr>")

                        html.append(f"<td>{ch}</td>")

                        for p in ps:

                            val = None

                            if src in structure[sec][grp][sg] and ch in structure[sec][grp][sg][src]:

                                val = structure[sec][grp][sg][src][ch].get(p)

                            # Check if this is a mismatch (value exists but differs)

                            is_mismatch = False

                            if val is not None:

                                # Check if there's a mismatch for this key

                                mismatch_key = (sec, grp, sg, p, ch)

                                is_mismatch = mismatch_key in mismatches

                            if val is None:

                                # No value - display "-"

                                html.append(f"<td class='missing'>-</td>")

                            elif is_mismatch:

                                # Mismatched value - highlight

                                html.append(f"<td class='bad'>{val}</td>")

                            else:

                                # Normal value

                                html.append(f"<td>{val}</td>")

                        html.append("</tr>")

                    html.append("</table>")

                html.append("</div>")

    html.append("</body></html>")

    return "".join(html)
 
def main():

    section_row = 5
 
    f1_path = Path(FILE1_FOLDER)

    f2_path = Path(FILE2_FOLDER)

    out_path = Path(OUTPUT_FOLDER)

    out_path.mkdir(parents=True, exist_ok=True)
 
    mapping = load_mapping()

    rpt_to_tc = reverse_mapping(mapping)
 
    file1_map = {}

    file2_map = {}
 
    for f in f1_path.glob("*.xlsx"):

        rpt_id = extract_rpt_id(f.name)

        if rpt_id and rpt_id in rpt_to_tc:

            tc = rpt_to_tc[rpt_id].replace("TC", "")

            file1_map[tc] = f
 
    for f in f2_path.glob("TC*_Exc_RBU*.xlsx"):

        tc = extract_tc(f.name)

        if tc:

            file2_map[tc] = f
 
    common_tcs = sorted(set(file1_map.keys()) & set(file2_map.keys()), key=int)
 
    print(f"Running {len(common_tcs)} scenarios...\n")
 
    for tc in common_tcs:

        file1 = file1_map[tc]

        file2 = file2_map[tc]
 
        g1 = read_xlsx(file1)

        g2 = read_xlsx(file2)
 
        r1 = extract_records(g1, file1.name, section_row)

        r2 = extract_records(g2, file2.name, section_row)
 
        data, mismatches = compare(r1 + r2)
 
        html = generate_html(data, mismatches, [file1.name, file2.name])
 
        output_file = out_path / f"TC{tc}_RBU_Report.html"
 
        with open(output_file, "w", encoding="utf-8") as f:

            f.write(html)
 
        print(f"TC{tc} completed")
 
    print("All scenarios completed.")
 
if __name__ == "__main__":

    main()