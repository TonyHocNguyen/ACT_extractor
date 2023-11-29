import pandas as pd
import time
import sys
import re
import os
import json

# Progress bar function (unchanged)
def progressBar(iterable, prefix='', suffix='', decimals=1, length=50, fill='█', printEnd='\r'):
    total = len(iterable)

    def printProgressBar(iteration):
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)

    printProgressBar(0)
    for i, item in enumerate(iterable):
        yield item
        printProgressBar(i + 1)
    print()

# Function to find VID (unchanged)
def find_VID(folder_path):
    pattern = "\w+"
    table_0 = pd.read_html(folder_path)[0]
    txt = table_0[1][1]
    VID = re.findall(pattern, txt)[2]
    return VID

# Function to extract ACT table (unchanged)
def extract_ACT(folder_path):
    table_6 = pd.read_html(folder_path)[6][1:]
    table_6.iloc[0, 0] = "Unit #"
    VID = find_VID(folder_path)
    table_6.iloc[:, 0] = VID
    return table_6

# Function to read folder paths from JSON (modified to escape backslashes)
def get_data_from_json(json_file):
    with open(json_file, 'r') as config_file:
        config_data = json.load(config_file)

    folder_paths = config_data.get("folder_paths", [])

    # Replace single backslashes with double backslashes
    folder_paths = ["/" + path.replace(os.sep, "/") for path in folder_paths]
    # folder_paths = [path.replace(os.sep, '\\\\') for path in folder_paths]
    
    folder_names = [os.path.basename(folder) for folder in folder_paths]

    return folder_paths, folder_names

# Function to join and save data (modified)
def join_data(dfs, folder_name, folder_path):
    try:
        df = pd.concat(dfs)
        df.columns = ["Unit #", "UUT Pin", "Functional Block",
                      "UUT Signal Name", "Package Net List", "Result Type"]

        file_name = f"{folder_name}-VIRAL.xlsx"
        file_path = os.path.join(folder_path, file_name)
        
        df.to_excel(file_path, index=False)

        print("ACT table extraction success!")
        return file_path
    
    except Exception as e:
        print("Excel file error!")
        print("Please check for any opened Excel file or available data in the VIRAL folder.")
        sys.exit(1)

# Main function (modified)
def main():
    if len(sys.argv) < 2:
        print("Usage: python script.py <config.json>")
        sys.exit(1)

    json_file = sys.argv[1]

    folder_paths, folder_names = get_data_from_json(json_file)
    file_paths = []

    for folder_path, folder_name in zip(folder_paths, folder_names):
        dfs = []

        try:
            items = os.listdir(folder_path)
            for item in progressBar(items, prefix=f'Progress for {folder_name}:', suffix='Complete', length=50):
                path = os.path.join(folder_path, item)
                if os.path.isdir(path):
                    for file in os.listdir(path):
                        if file.endswith(".html"):
                            table = extract_ACT(os.path.join(path, file))
                            dfs.append(table)
                time.sleep(0.1)
        except Exception as e:
            print(f"Path error for {folder_name}!")
            print("Please check for VPN connection or correct VIRAL folder path.")
            sys.exit(1)

        file_path = join_data(dfs, folder_name, folder_path)
        file_paths.append(file_path)
        # print(f"{folder_name} output file path: {file_path}")

    return folder_paths, folder_names, file_paths

if __name__ == "__main__":
    folder_paths, folder_names, file_paths = main()
