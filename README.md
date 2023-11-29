**ACT extractor user guide**\
The script will loop through all the sub-folders in the directory and extract the ACT table from the VIRAL html files.\
The script will then concatenate all unit data, create a pivot table, and save it to an Excel file. The Excel file will be saved to the shared drive VIRAL folder of the CI case

**General library requirements**\
Please follow this step carefully before running the script.
The script is developed in Python 3.10.
Please check for compatibility of the packages with your Python version in `requirements.txt`.

To download the required packages and libraries:
1. Connect to a different network from the Intel network or disconnect from the VPN.
2. Open the terminal (command prompt/cmd) and change the directory to the folder containing the script.
3. Locate where Python 3.10 is installed on your PC using `where python` in the terminal
4. Copy the python path, edit the `install.bat` file and change the python path accordingly.
5. Run the `install.bat` file to install.

**Script instructions**\
These are the steps to run the script once you have installed the required libraries.

Add cases for extraction:
1. In the `config.json` file, add the VIRAL folder names from the FACR share drive (with FACR case number and product name) and save the config file.
2. Put the cases as a string (include `""` for each case and separate with `,` if there are more than 1 case) according to the year.

Example:
```
{
    "VIRAL":
    [
        {
            "year": 2023,
            "cases": ["CI2334-5660 RPL-S881 BGA", "CI23XX-XXXX ADL-P682"]
        }
    ]
}
```

Run the ACT extractor script using the .bat file:
1. Connect to the Intel VPN or Intel network.
2. Double click the `act_extractor.bat`, this will open a terminal where you can view the progress (run by administrator if it fails to open a terminal)

If the script raises any path or file error:
1. Check for VPN / Intel network connection.
2. Check for correct folder name or path input in the `config.json` file. 
3. Close any opened Excel file before running the script again.

If you want to abort the extraction, press `Ctrl + C` or close the terminal.