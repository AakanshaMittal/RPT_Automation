import requests
from pathlib import Path
import time
 
BASE_URL = "https://actuarialhub-dev-api.optum.com/api/azure-blob/download"
TOKEN_FILE = Path(__file__).resolve().parent.parent / "token.txt"


def get_headers():
    token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    return {"Authorization": f"Bearer {token}"}


def download_all_files(tab, mapping, base_data_path):
    headers = get_headers()
    save_folder = Path(base_data_path) / tab / "File1"
    save_folder.mkdir(parents=True, exist_ok=True)
    print(f"Downloading files for {tab}...")
    for tc, rpt_id in mapping.items():
 
        url = f"{BASE_URL}/{tab}/{rpt_id}"
        file_name = f"RPT_{rpt_id}_{tab}.xlsx"
        file_path = save_folder / file_name
 
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                with open(file_path, "wb") as f:
                    f.write(response.content)
                print(f"{tc} downloaded -> {file_name}")
 
            else:
                print(f"{tc} failed (status {response.status_code})")
 
        except Exception as e:

            print(f"{tc} failed -> {e}")
 