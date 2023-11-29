@echo off

:: Specify the virtual environment path
set VENV_PATH=%USERPROFILE%\venv

:: Specify the Python script and configuration file
set PYTHON_SCRIPT=main.py
set CONFIG_FILE=config.json

:: Delete the gen_py folder in the temp directory
rd /s /q %temp%\gen_py

:: Run the extraction script using the virtual environment
%VENV_PATH%\Scripts\python.exe %PYTHON_SCRIPT% %CONFIG_FILE%

pause