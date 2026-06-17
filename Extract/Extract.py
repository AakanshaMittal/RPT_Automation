import pandas as pd
import os

# ===== CONFIG =====
file_path = "RPTFiles/T20_rebate_refresh_oglp.xlsm"
output_dir = "TabsFiles"
tc_code = "TC1"   # <-- You can dynamically pass this

# Only extract these sheets (Requirement 2)
target_sheets = [
    "Detailed Stats (Internal Use2)",
    "RatesPerQnty",
    "RA Output",
    "Rate Build-up",
    "FinAlign"
]

os.makedirs(output_dir, exist_ok=True)

# ===== LOAD FILE (WITHOUT RUNNING MACROS) =====
xls = pd.ExcelFile(file_path, engine="openpyxl")

# ===== PROCESS ONLY REQUIRED SHEETS =====
for sheet in target_sheets:
    if sheet not in xls.sheet_names:
        print(f"[WARNING] Sheet not found: {sheet}")
        continue

    df = xls.parse(sheet)

    # ===== CLEAN TAB NAME FOR FILE =====
    clean_sheet_name = sheet.replace(" ", "").replace("/", "").replace("-", "")

    # ===== CREATE FILE NAME (Requirement 1) =====
    output_file = f"{tc_code}_Exc_{clean_sheet_name}.xlsx"

    df.to_excel(os.path.join(output_dir, output_file), index=False)

    print(f"[DONE] Extracted -> {output_file}")