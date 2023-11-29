"""
ACT extractor v2.0
UNOFFICIAL RELEASE - For PQE team internal use only

UPDATE v1.0: basic extraction and concatenation functions
UPDATE v1.1: added progress bar, user input to access share drive, and error catch for extract operation
UPDATE v1.2: added check for CI case year from user input. Maximum year range 2000-2099
UPDATE v1.3: removed repeated header, changed input to folder link
UPDATE v2.0: added pivot table creation functionality

Please read the detailed instructions in README.md
"""
__author__ = "Nguyen, Hoc"
__email__ = "hoc.nguyen@intel.com"
__version__ = "2.0"
__status__ = "Development"
__date__ = "04 October 2023"

import pandas as pd
import time
import sys
import re
import os
import json
from tqdm import tqdm  # Import tqdm for progress bars

# Function to find VID (unchanged)
def find_VID(file_path):
    pattern = "\w+"
    table_0 = pd.read_html(file_path)[0]
    txt = table_0[1][1]
    VID = re.findall(pattern, txt)[2]
    return VID

# Function to extract ACT table (unchanged)
def extract_ACT(file_path):
    table_6 = pd.read_html(file_path)[6][1:]
    table_6.iloc[0, 0] = "Unit #"
    VID = find_VID(file_path)
    table_6.iloc[:, 0] = VID
    return table_6

# Function to read folder paths from JSON (modified to use base path)
def get_data_from_json(json_file, DEBUG=False):
    with open(json_file, 'r') as config_file:
        config_data = json.load(config_file)

    if DEBUG:
        # Only use for debugs
        PATH = config_data.get("DEBUG_PATH")
    else:
        # Use base path to construct folder paths
        PATH = config_data.get("BASE_PATH")

    VIRALS = config_data.get("VIRAL", [])

    case_paths = []
    for viral in VIRALS:
        year = viral.get("year")
        cases = viral.get("cases")

        # Construct the folder path using the base path, year, and cases
        year_path = os.path.join(PATH, str(year))
        for case in cases:
            case_path = os.path.join(year_path, case)
            case_paths.append(case_path)

    case_names = [os.path.basename(folder) for folder in case_paths]
    
    return case_paths, case_names

# Function to join and save data (unchanged)
def join_data(dfs, case_name, case_path):
    try:
        df = pd.concat(dfs)
        df.columns = ["Unit #", "UUT Pin", "Functional Block",
                      "UUT Signal Name", "Package Net List", "Result Type"]

        excel_name = f"{case_name}-VIRAL.xlsx"
        excel_path = os.path.join(case_path, excel_name)
        
        df.to_excel(excel_path, index=False)

        return excel_path
    
    except Exception as e:
        print("Excel file error!")
        print("Please check for any opened Excel file or available data in the VIRAL folder.")
        sys.exit(1)

# Main function (modified for base path)
def run_act(DEBUG=False):
    if len(sys.argv) < 2:
        print("Usage: python script.py <config.json>")
        sys.exit(1)

    json_file = sys.argv[1]

    if DEBUG:
        case_paths, case_names = get_data_from_json(json_file, DEBUG=True)
    else:
        case_paths, case_names = get_data_from_json(json_file)
    
    excel_paths = []

    for case_path, case_name in zip(case_paths, case_names):
        dfs = []

        try:
            items = os.listdir(case_path)
            for item in tqdm(items, desc=f'Progress for {case_name}', ncols=100):
                path = os.path.join(case_path, item)
                if os.path.isdir(path):
                    for file in os.listdir(path):
                        if file.endswith('.html'):
                            table = extract_ACT(os.path.join(path, file))
                            dfs.append(table)
                time.sleep(0.1)
        except Exception as e:
            print(f"Path error for {case_name}!")
            print("Please check for VPN connection or correct VIRAL folder path.")
            sys.exit(1)

        excel_path = join_data(dfs, case_name, case_path)
        excel_paths.append(excel_path)

    return excel_paths

def test():
    run_act(DEBUG=True)

if __name__ == "__main__":
    test()
