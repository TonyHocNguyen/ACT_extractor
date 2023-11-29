#!/usr/bin/env python3
"""
ACT extractor v1.3
UNOFFICIAL RELEASE - For PQE team internal use only

UPDATE v1.0: basic extraction and concatenation functions
UPDATE v1.1: added progress bar, user input to access share drive, and error catch for extract operation
UPDATE v1.2: added check for CI case year from user input. Maximum year range 2000-2099
UPDATE v1.3: removed repeated header, changed input to folder link

Please read the detailed instructions in README.md
"""
__author__ = "Nguyen, Hoc"
__email__ = "hoc.nguyen@intel.com"
__version__ = "1.3"
__status__ = "Development"
__date__ = "21 August 2023"

import pandas as pd
import time
import sys
import re
import os

def progressBar(iterable, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iterable    - Required  : iterable object (Iterable)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    total = len(iterable)
    # Progress Bar Printing Function
    def printProgressBar (iteration):
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Initial Call
    printProgressBar(0)
    # Update Progress Bar
    for i, item in enumerate(iterable):
        yield item
        printProgressBar(i + 1)
    # Print New Line on Complete
    print()

def find_VID(URL):
    # Check for VID pattern in the text
    pattern = "\w+"
    table_0 = pd.read_html(URL)[0]
    txt = table_0[1][1]
    VID = re.findall(pattern, txt)[2]

    return VID

def extract_ACT(URL): 
    # Assign the table data to a Pandas dataframe
    table_6 = pd.read_html(URL)[6][1:]
    # Change the first column name to unit #
    table_6[0][0] = "Unit #"
    # Extract Visual ID from the file
    VID = find_VID(URL)
    # Replace column values to VID
    table_6[0][:] = VID
    
    return table_6

def get_data():
    # Get the input for VIRAL link
    if len(sys.argv) == 1:
        URL = input("Please enter the link to your VIRAL folder: ")
        URL = URL.replace(os.sep, '/')
        # Get the folder name
        folder_name = os.path.basename(URL)
    else:
        URL = ' '.join(sys.argv[1:])
        URL = URL.replace(os.sep, '/')
        # Get the folder name
        folder_name = os.path.basename(URL)

    # A list of all the tables
    dfs = []

    try:
        # Look for items in the folder
        items = os.listdir(URL)
        # A Nicer, Single-Call Usage
        for item in progressBar(items, prefix = 'Progress:', suffix = 'Complete', length = 50):
            # Do stuff...
            path = os.path.join(URL, item)
            if os.path.isdir(path):
                # Loop through files in the sub folder
                for file in os.listdir(URL + "/" + item):
                    # Check if the file is a html file
                    if file.endswith(".html"):
                        # Extract the table from the html file
                        table = extract_ACT(URL + "/" + item + "/" + file)
                        # Add the table to the dataframe
                        dfs.append(table)
            time.sleep(0.1)
            continue

    except Exception as e:
        print("Path error! Please check for VPN connection or correct VIRAL folder path.")

    return dfs, folder_name, URL

def join_data(dfs, folder_name, URL):
    # catch concatenate error
    try:
        # Concatenate all the tables
        df = pd.concat(dfs)
        
        # Add column names
        df.columns = ["Unit #", "UUT Pin", "Functional Block", "UUT Signal Name", "Package Net List", "Result Type"]
        
        # Dataframe to excel
        # Save to local folder
        df.to_excel(folder_name + "-VIRAL.xlsx", index=False)
        
        # Save to shared VIRAL folder
        file_path = URL + "/" + folder_name + "-VIRAL.xlsx"
        df.to_excel(file_path, index=False)
        print("ACT table extraction success!")
        return file_path
        
    except Exception as e:
        print("File error! Please check for any opened Excel file or available data in the VIRAL folder.")

def main():
    dfs, folder_name, URL = get_data()
    # file_path = join_data(dfs, folder_name, URL)
    print(URL)
      
if __name__ == "__main__":
    main()