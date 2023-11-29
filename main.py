import act_extractor_v2
import create_pivot_table
import time
import os

def main():
    # Extract data from paths in json config and combine data into excel files
    print("Extracting ACT data for cases...")
    file_paths = act_extractor_v2.run_act(DEBUG=False)

    # Creating pivot tables in the excel files
    for file_path in file_paths:
        folder_name = os.path.basename(file_path)
        print(f'\nCreating pivot table for {folder_name}')
        create_pivot_table.run_excel(file_path)
        print('Pivot table creation success at:')
        print(file_path)
        time.sleep(0.1)

if __name__ == '__main__':
    main()