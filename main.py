import json
import importlib
from utils.downloader import download_all_files
 
CONFIG_PATH = "config/mapping.json"
BASE_DATA_PATH = "data"
 
def load_mapping():

    with open(CONFIG_PATH) as f:
        return json.load(f)
 
 
def run_tab(tab):

    mapping = load_mapping()
    download_all_files(tab, mapping, BASE_DATA_PATH)
    module_name = f"comparator.{tab}"
    comparator_module = importlib.import_module(module_name)
    print(f"\nRunning comparator for {tab}...\n")
    comparator_module.main()
 
 
if __name__ == "__main__":
 #  
    tabs = [ "internalStats2", "externalStats", "ratesPerQuantity", "raOutput", "rateBuildUp", "fmsForecastOutput", "finAlign",  "outputSpecialityRpt",]
 
    for tab in tabs:
        run_tab(tab)