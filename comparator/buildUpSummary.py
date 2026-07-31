import os
import re
import json
from pathlib import Path

import pandas as pd

# =========================================================
# CONFIGURATION (ALIGNED WITH RPQ FRAMEWORK)
# =========================================================

BASE_PATH = Path("data/buildUpSummary")

FILE1_FOLDER = BASE_PATH / "File1"
FILE2_FOLDER = BASE_PATH / "File2"
OUTPUT_ROOT = BASE_PATH / "Output"

mapping_file = r"config/mapping.json"

TOLERANCE = 0.0001


# =========================================================
# MAPPING METHODS
# =========================================================

def load_mapping():
    with open(mapping_file) as f:
        return json.load(f)


def reverse_mapping(mapping):
    return {v: k for k, v in mapping.items()}


def extract_rpt_id(filename):
    match = re.search(r"(RPT\d+[A-Z]{2})", filename.upper())
    return match.group(1) if match else None


def extract_tc(filename):
    match = re.search(r"TC(\d+[A-Z]?)", filename.upper())
    return match.group(1) if match else None


# =========================================================
# FILE READING
# =========================================================

def read_excel_safe(path):
    return (
        pd.read_excel(path, header=None, engine="openpyxl")
        .astype(object)
        .fillna("")
    )


# =========================================================
# NORMALIZATION & MATCHING (UNCHANGED)
# =========================================================

def normalize(val):
    if val in ("", None) or pd.isna(val):
        return None

    val = str(val).strip().replace(",", "")

    if val.endswith("%"):
        try:
            return float(val.replace("%", "")) / 100
        except Exception:
            return val

    try:
        return float(val)
    except Exception:
        return val.lower()


def is_match(v1, v2):
    n1 = normalize(v1)
    n2 = normalize(v2)

    if isinstance(n1, float) and isinstance(n2, float):
        return abs(n1 - n2) <= TOLERANCE

    return n1 == n2


# =========================================================
# HTML REPORT + MISMATCH COUNT
# =========================================================

def compare_files(file1_path, file2_path, output_dir, report_name):

    file1_df = read_excel_safe(file1_path)
    file2_df = read_excel_safe(file2_path)

    max_rows = max(len(file1_df), len(file2_df))
    max_cols = max(file1_df.shape[1], file2_df.shape[1])

    mismatch_count = 0

    html = f"""
    <html>
    <head>
        <title>{report_name}</title>
        <style>
            body {{ font-family: Calibri, Arial, sans-serif; }}
            table {{ border-collapse: collapse; width: 100%; }}
            td {{ border: 1px solid #333; padding: 6px; white-space: pre-wrap; }}
            .mismatch {{ background-color: #ffcccc; }}
            .header {{ font-weight: bold; background-color: #d9d9d9; }}
        </style>
    </head>
    <body>
        <h2>Comparison Report : {report_name}</h2>
        <p><b>Format:</b> File1 value | File2 value</p>
        <table>
    """

    for r in range(max_rows):

        html += "<tr>"

        for c in range(max_cols):

            file1_val = (
                file1_df.iat[r, c]
                if r < len(file1_df) and c < file1_df.shape[1]
                else ""
            )

            file2_val = (
                file2_df.iat[r, c]
                if r < len(file2_df) and c < file2_df.shape[1]
                else ""
            )

            mismatch = not is_match(file1_val, file2_val)

            if mismatch:
                mismatch_count += 1

            css_class = "mismatch" if mismatch else ""

            if r == 0:
                css_class += " header"

            html += (
                f'<td class="{css_class}">'
                f'{file1_val} | {file2_val}'
                f'</td>'
            )

        html += "</tr>"

    html += """
        </table>
    </body>
    </html>
    """

    os.makedirs(output_dir, exist_ok=True)

    with open(
        os.path.join(output_dir, f"{report_name}.html"),
        "w",
        encoding="utf-8"
    ) as f:
        f.write(html)

    return mismatch_count


# =========================================================
# CONSOLIDATED REPORT
# =========================================================

def generate_consolidated_report(data, output_path):

    df = pd.DataFrame(data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

    print(f"✅ Consolidated report generated: {output_path}")


# =========================================================
# MAIN
# =========================================================

def main():

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    mapping = load_mapping()
    rpt_to_tc = reverse_mapping(mapping)

    file1_map = {}
    file2_map = {}

    # -----------------------------------------------------
    # FILE1 (RPT BASED)
    # -----------------------------------------------------

    for f in FILE1_FOLDER.glob("*.xlsx"):

        rpt_id = extract_rpt_id(f.name)

        if rpt_id and rpt_id in rpt_to_tc:
            file1_map[
                rpt_to_tc[rpt_id].replace("TC", "")
            ] = f

    # -----------------------------------------------------
    # FILE2 (TC BASED)
    # -----------------------------------------------------

    for f in FILE2_FOLDER.glob("*.xlsx"):

        tc = extract_tc(f.name)

        if tc:
            file2_map[tc] = f

    common_tcs = sorted(
        set(file1_map) & set(file2_map),
        key=str
    )

    print("FILE1 TCs:", sorted(file1_map.keys()))
    print("FILE2 TCs:", sorted(file2_map.keys()))
    print("COMMON TCs:", common_tcs)

    if not common_tcs:
        print("❌ No matching TC files found.")
        return

    consolidated_data = []

    for tc in common_tcs:

        file1 = file1_map[tc]
        file2 = file2_map[tc]

        output_dir = OUTPUT_ROOT / f"TC{tc}"

        mismatch_count = compare_files(
            str(file1),
            str(file2),
            str(output_dir),
            f"TC{tc}"
        )

        consolidated_data.append(
            {
                "Test Case": f"TC{tc}",
                "Mismatch Count": mismatch_count,
                "Status": (
                    "PASS"
                    if mismatch_count == 0
                    else "FAIL"
                )
            }
        )

    consolidated_file = (
        OUTPUT_ROOT /
        "Consolidated_BuildUp_Report.xlsx"
    )

    generate_consolidated_report(
        consolidated_data,
        str(consolidated_file)
    )

    print("✅ All scenarios completed.")


if __name__ == "__main__":
    main()