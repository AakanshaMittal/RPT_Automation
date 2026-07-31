import os
import io
import re
import requests
import pandas as pd

# =========================================================
# CONFIGURATION
# =========================================================

API_BASE_URL = "https://actuarialhub-dev-api.optum.com/api/azure-blob/download/rptForecastOutputPl/"

API_ENDPOINTS = {
     5: "RPT6072999484647YM",
     19: "RPT3354366772570ZF"
    # 19: "RPT6936998921199VX"
# 7: "RPT5778326415391WR",
# 16: "RPT4916668589193MM",
# 12: "RPT6005053739424UO",
# 20: "RPT6369172094038YQ",
# 4: "RPT5102510432402PT",
# 11: "RPT9352764916055DM",
# 17: "RPT614783834725AQ"
#     1:"RPT8144866384611ZC",
# 2:"RPT5621045155116MQ",
# 3 : "RPT6001956739171YP",
# 4 : "RPT6675435528517RS",
# 5 : "RPT8425438558663AN",
# 6 : "RPT1676176927296BS",
# 7 : "RPT5778326415391WR",
# 8 : "RPT2602083498305GG",
# 9 : "RPT2611325815634MP",
# 10 : "RPT8077467244232BD",
# 11 : "RPT9352764916055DM",
# 12 : "RPT3569087553929AT",
# 13 : "RPT1511618462593QI",
# 14: "RPT7742757019314VB",
# 15 : "RPT9135597151051ZB",
# 9: "RPT2611325815634MP",
# 15: "RPT4238213741826PC",
# 16: "RPT2579918051352BZ"
# 16 : "RPT6052896907489SK",
# 17: "RPT5063627014454ZD",
# 18 : "RPT7645649907263TU",
# 19 : "RPT3816552382848TK",
# 20 : "RPT6369172094038YQ"
    # 17: "RPT9518641008898FM"
    # 20: "RPT1958363730057DJ"
    # 4: "RPT6610742477102RP"
    # 11: "RPT4044886397592RW"
    #  21: "RPT8984896822403SS",
    #  7: "RPT6348854143852QW"
    # 4: "RPT5321116414285FC"
    # 1: "RPT8028197735728WS",
# 2: "RPT6173408222224NJ",
# 3: "RPT6430463565288YY",
# 4: "RPT9093292786022KE",
# 5: "RPT1404880669668VS",
# 6: "RPT6412645171166EW",
# 7: "RPT5340512932842BY",
# 8: "RPT8967267773183KS",
# 9: "RPT1671330770979CF",
# 10: "RPT8107702707903CZ",
# 11: "RPT6773589255724NR",
# 12: "RPT3148035460123ML",
# 13: "RPT3450970068231MO",
# 14: "RPT6649501860139LC",
# 15: "RPT7248685682065JU",
# 16: "RPT6154188535123LI",
# 17: "RPT6724285789423RU",
# 18: "RPT7166282902057JT",
# 19: "RPT7014190005476SB",
# 20: "RPT5738801382770ZX"
# 21: "RPT4518612165658NM"
}

FOLDER_APP = "./ForecastPL_APP_Folder_Dev_July8"
FOLDER_EXC = "./ForecastPL_EXC_RebateRefresh"
OUTPUT_ROOT = "./comparison_reports_RebateRefresh_July8"

TOLERANCE = 0.0001

API_HEADERS = {}

# =========================================================
# SAFE EXCEL READ (UNCHANGED)
# =========================================================
def read_excel_safe(source):
    if isinstance(source, bytes):
        df = pd.read_excel(io.BytesIO(source), header=None, engine="openpyxl")
    else:
        df = pd.read_excel(source, header=None, engine="openpyxl")

    return df.astype(object).fillna("")

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
        except:
            return val

    try:
        return float(val)
    except:
        return val.lower()

def is_match(v1, v2):
    n1 = normalize(v1)
    n2 = normalize(v2)

    if isinstance(n1, float) and isinstance(n2, float):
        return abs(n1 - n2) <= TOLERANCE

    return n1 == n2

# =========================================================
# DOWNLOAD APP FILES FROM API
# =========================================================
def download_app_files():
    os.makedirs(FOLDER_APP, exist_ok=True)
    downloaded = {}

    for file_no in sorted(API_ENDPOINTS.keys()):
        endpoint = API_ENDPOINTS[file_no]

        url = f"{API_BASE_URL}{endpoint}"
        response = requests.get(url, headers=API_HEADERS)

        if response.status_code != 200:
            raise RuntimeError(f"Failed to download APP file T{file_no}")

        file_path = os.path.join(
            FOLDER_APP, f"T{file_no}_App_ForecastPL.xlsx"
        )

        with open(file_path, "wb") as f:
            f.write(response.content)

        downloaded[file_no] = file_path

    return downloaded

# =========================================================
# MAP EXC FILES
# =========================================================
def extract_t_number(filename):
    match = re.search(r"T(\d+)_Exc_ForecastPL", filename, re.IGNORECASE)
    return int(match.group(1)) if match else None

def get_exc_files():
    exc_files = {}

    for file in os.listdir(FOLDER_EXC):
        if file.lower().endswith(".xlsx"):
            t_no = extract_t_number(file)
            if t_no is not None:
                exc_files[t_no] = os.path.join(FOLDER_EXC, file)

    return exc_files

# =========================================================
# HTML REPORT GENERATION + MISMATCH COUNT (Minor Addition)
# =========================================================
def compare_files(app_path, exc_path, output_dir, report_name):

    app_df = read_excel_safe(app_path)
    exc_df = read_excel_safe(exc_path)

    max_rows = max(len(app_df), len(exc_df))
    max_cols = max(app_df.shape[1], exc_df.shape[1])

    mismatch_count = 0   # ⭐ NEW

    html = f"""
    <html>
    <head>
        <title>{report_name}</title>
        <style>
            body {{ font-family: Calibri, Arial; }}
            table {{ border-collapse: collapse; width: 100%; }}
            td {{ border: 1px solid #333; padding: 6px; }}
            .mismatch {{ background-color: #ffcccc; }}
        </style>
    </head>
    <body>
        <h2>{report_name}</h2>
        <table>
    """

    for r in range(max_rows):
        html += "<tr>"
        for c in range(max_cols):
            app_val = app_df.iat[r, c] if r < len(app_df) and c < app_df.shape[1] else ""
            exc_val = exc_df.iat[r, c] if r < len(exc_df) and c < exc_df.shape[1] else ""

            is_mis = not is_match(app_val, exc_val)

            if is_mis:
                mismatch_count += 1   # ⭐ COUNT

            css = "mismatch" if is_mis else ""
            html += f'<td class="{css}">{app_val} | {exc_val}</td>'

        html += "</tr>"

    html += "</table></body></html>"

    os.makedirs(output_dir, exist_ok=True)
    with open(os.path.join(output_dir, f"{report_name}.html"), "w", encoding="utf-8") as f:
        f.write(html)

    return mismatch_count   # ⭐ RETURN

# =========================================================
# CONSOLIDATED EXCEL REPORT (NEW)
# =========================================================
def generate_consolidated_excel(results, path):
    df = pd.DataFrame(results)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")

    print(f"✅ Consolidated report: {path}")

# =========================================================
# MAIN EXECUTION
# =========================================================
if __name__ == "__main__":

    os.makedirs(OUTPUT_ROOT, exist_ok=True)

    app_files = download_app_files()
    exc_files = get_exc_files()

    consolidated_results = []   # ⭐ NEW

    for t_no in sorted(app_files.keys()):
        if t_no not in exc_files:
            continue

        report_name = f"T{t_no}_Report_ForecastPL"
        report_dir = os.path.join(OUTPUT_ROOT, f"T{t_no}")

        mismatch_count = compare_files(
            app_files[t_no],
            exc_files[t_no],
            report_dir,
            report_name
        )

        # ⭐ STORE SUMMARY
        consolidated_results.append({
            "Test Case": f"T{t_no}",
            "Mismatch Count": mismatch_count,
            "Status": "PASS" if mismatch_count == 0 else "FAIL"
        })

    # ⭐ GENERATE EXCEL
    consolidated_path = os.path.join(OUTPUT_ROOT, "Consolidated_Report.xlsx")
    generate_consolidated_excel(consolidated_results, consolidated_path)

    print("✅ Detailed ForecastPL comparison reports generated successfully.")