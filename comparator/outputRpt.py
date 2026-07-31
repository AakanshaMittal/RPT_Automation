import re
import json
from pathlib import Path

import pandas as pd

# =====================================================
# CONFIGURATION
# =====================================================

BASE_PATH = Path("data/outputRpt")

FILE1_FOLDER = BASE_PATH / "File1"
FILE2_FOLDER = BASE_PATH / "File2"
OUTPUT_FOLDER = BASE_PATH / "Output"

mapping_file = r"config/mapping.json"

TOLERANCE = 0.0001


# =====================================================
# MAPPING METHODS
# =====================================================

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


# =====================================================
# ORIGINAL COMPARISON LOGIC
# =====================================================

def load_excel(path):

    df = pd.read_excel(path, header=3)

    df.dropna(how="all", inplace=True)

    col_map = {
        df.columns("Key"),
        df.columns("Value")
    }

    df = df.rename(columns=col_map)

    df = df[["Key", "Value"]]

    df = df[df["Key"].notna()]
    df = df[df["Key"].astype(str).str.strip() != ""]

    df["Value"] = df["Value"].fillna("")

    return df


def compare_excels(file1, file2, report_path):

    df1 = load_excel(file1)
    df2 = load_excel(file2)

    merged = pd.merge(
        df1,
        df2,
        on="Key",
        how="inner",
        suffixes=("_File1", "_File2")
    )

    merged["Value_File1_num"] = pd.to_numeric(
        merged["Value_File1"],
        errors="coerce"
    )

    merged["Value_File2_num"] = pd.to_numeric(
        merged["Value_File2"],
        errors="coerce"
    )

    merged["Diff"] = (
        merged["Value_File1_num"] -
        merged["Value_File2_num"]
    ).abs()

    mismatches = merged[
        (merged["Diff"] >= TOLERANCE)
        |
        (
            (merged["Value_File1"] != merged["Value_File2"])
            &
            (
                merged["Value_File1_num"].isna()
                |
                merged["Value_File2_num"].isna()
            )
        )
    ].copy()

    mismatches["Value_File1"] = mismatches["Value_File1"].fillna("")
    mismatches["Value_File2"] = mismatches["Value_File2"].fillna("")
    mismatches["Diff"] = mismatches["Diff"].fillna("")

    def normalize(v):
        try:
            if float(v) == 0:
                return "0"
        except Exception:
            pass

        return v

    mismatches["Value_File1"] = mismatches["Value_File1"].apply(normalize)
    mismatches["Value_File2"] = mismatches["Value_File2"].apply(normalize)

    summary = {
        "File 1 Path": str(Path(file1).resolve()),
        "File 2 Path": str(Path(file2).resolve()),
        "Rows in File 1": len(df1),
        "Rows in File 2": len(df2),
        "Common Keys Compared": len(merged),
        "Value Mismatches (>= 0.0001)": len(mismatches)
    }

    html = """
<html>
<head>
<title>Excel Comparison Report</title>
<style>
body { font-family: Arial; padding:20px; }
table { border-collapse: collapse; width:100%; margin-top:20px; }
th, td { border:1px solid #555; padding:8px; text-align:left; }
th { background-color:#eee; }
h2 { color:#003366; }
</style>
</head>
<body>

<h1>Excel Comparison Report</h1>

<h2>Summary</h2>

<table>
"""

    for k, v in summary.items():
        html += f"<tr><th>{k}</th><td>{v}</td></tr>"

    html += "</table>"

    html += "<h2>Value Mismatches (Difference >= 0.0001)</h2>"

    if mismatches.empty:

        html += "<p>No mismatches found above threshold.</p>"

    else:

        html += mismatches[
            ["Key", "Value_File1", "Value_File2", "Diff"]
        ].to_html(index=False)

    html += """
</body>
</html>
"""

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ Report generated: {report_path}")

    return len(mismatches)


# =====================================================
# CONSOLIDATED REPORT
# =====================================================

def generate_consolidated_report(data, output_path):

    df = pd.DataFrame(data)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

    print(
        f"✅ Consolidated report generated: "
        f"{output_path}"
    )


# =====================================================
# MAIN
# =====================================================

def main():

    OUTPUT_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    mapping = load_mapping()

    rpt_to_tc = reverse_mapping(mapping)

    file1_map = {}
    file2_map = {}

    # ----------------------------------------
    # FILE1 (APP / RPT FILES)
    # ----------------------------------------

    for f in FILE1_FOLDER.glob("*.xlsx"):

        rpt_id = extract_rpt_id(f.name)

        if rpt_id and rpt_id in rpt_to_tc:

            file1_map[
                rpt_to_tc[rpt_id].replace("TC", "")
            ] = f

    # ----------------------------------------
    # FILE2 (EXC / TC FILES)
    # ----------------------------------------

    for f in FILE2_FOLDER.glob("*.xlsx"):

        tc = extract_tc(f.name)

        if tc:
            file2_map[tc] = f

    common_tcs = sorted(
        set(file1_map) & set(file2_map),
        key=str
    )

    print("FILE1 TCs :", sorted(file1_map.keys()))
    print("FILE2 TCs :", sorted(file2_map.keys()))
    print("COMMON TCs:", common_tcs)

    if not common_tcs:
        print("❌ No matching TC files found.")
        return

    consolidated_data = []

    for tc in common_tcs:

        try:

            print(f"\nProcessing TC{tc}")

            file1 = file1_map[tc]
            file2 = file2_map[tc]

            report_file = (
                OUTPUT_FOLDER /
                f"TC{tc}_OutputRpt_Report.html"
            )

            mismatch_count = compare_excels(
                str(file1),
                str(file2),
                str(report_file)
            )

            consolidated_data.append({
                "Test Case": f"TC{tc}",
                "Mismatch Count": mismatch_count,
                "Status":
                    "PASS"
                    if mismatch_count == 0
                    else "FAIL"
            })

        except Exception as e:

            print(f"❌ TC{tc} failed: {e}")

            consolidated_data.append({
                "Test Case": f"TC{tc}",
                "Mismatch Count": "ERROR",
                "Status": "ERROR"
            })

    consolidated_file = (
        OUTPUT_FOLDER /
        "Consolidated_OutputRpt_Report.xlsx"
    )

    generate_consolidated_report(
        consolidated_data,
        str(consolidated_file)
    )

    print("✅ All scenarios completed.")


if __name__ == "__main__":
    main()