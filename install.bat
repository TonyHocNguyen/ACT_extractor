@echo off

:: Specify the Python interpreter path
set PYTHON_PATH=C:\Python310\python.exe

:: Specify the virtual environment path
set "VENV_PATH=%USERPROFILE%\venv"

:: Specify the requirements file
set "REQUIREMENTS_FILE=requirements.txt"

:: Step 0: Delete the existing virtual environment (if it exists)
if exist "%VENV_PATH%" (
    echo Deleting existing virtual environment...
    rmdir /s /q "%VENV_PATH%"
)

:: Step 1: Create the new virtual environment
echo Creating virtual environment...
"%PYTHON_PATH%" -m venv "%VENV_PATH%"

:: Step 2: Activate the virtual environment
echo Activating virtual environment...
call "%VENV_PATH%\Scripts\activate"

:: Debugging: Print the active Python interpreter to verify it's the one in the virtual environment.
echo Active Python interpreter: %PATH%

:: Step 3: Install the required packages from requirements.txt
echo Installing packages...
pip install -r "%REQUIREMENTS_FILE%" || (
    echo Package installation failed or timed out.
    echo Please check for errors and network issues.
    goto :EOF
)

:: Step 4: Deactivate the virtual environment
echo Deactivating virtual environment...
deactivate

:: Indicate the completion of the installation
echo Installation completed.

:: Pause for a moment before closing the command prompt
pause
