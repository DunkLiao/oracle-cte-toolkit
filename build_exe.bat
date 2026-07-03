@echo off
setlocal

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "VENV_DIR=%ROOT%\build\package-venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "OUT_DIR=%ROOT%\dist\oracle-cte-toolkit"
set "WORK_DIR=%ROOT%\build\pyinstaller"

cd /d "%ROOT%"

echo [1/6] Checking system Python...
python --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found. Please install Python 3.9+.
    exit /b 1
)

echo [2/6] Preparing isolated build virtual environment...
if not exist "%PYTHON_EXE%" (
    python -m venv "%VENV_DIR%"
    if errorlevel 1 exit /b 1
)

echo [3/6] Installing build dependencies...
if /i "%SKIP_PIP_INSTALL%"=="1" (
    echo SKIP_PIP_INSTALL=1, skipping pip install.
) else (
    "%PYTHON_EXE%" -m pip install --upgrade pip
    if errorlevel 1 exit /b 1
    "%PYTHON_EXE%" -m pip install --upgrade pyinstaller oracledb pandas openpyxl chardet
    if errorlevel 1 exit /b 1
)

echo [4/6] Checking dependencies...
"%PYTHON_EXE%" -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: PyInstaller is not installed in the build virtual environment.
    exit /b 1
)

"%PYTHON_EXE%" -c "import oracledb, pandas, openpyxl, chardet" >nul 2>nul
if errorlevel 1 (
    echo ERROR: Missing one or more application dependencies: oracledb pandas openpyxl chardet
    exit /b 1
)

echo [5/6] Cleaning previous package output...
if exist "%OUT_DIR%" rmdir /s /q "%OUT_DIR%"
if exist "%WORK_DIR%" rmdir /s /q "%WORK_DIR%"
mkdir "%OUT_DIR%"

echo [6/6] Building OracleQueryTool.exe...
"%PYTHON_EXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name OracleQueryTool ^
    --distpath "%OUT_DIR%" ^
    --workpath "%WORK_DIR%" ^
    --specpath "%WORK_DIR%" ^
    --hidden-import config_loader ^
    --hidden-import oracle_client ^
    --hidden-import sql_reader ^
    --hidden-import excel_exporter ^
    --hidden-import chardet ^
    --hidden-import oracledb ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    "%ROOT%\oracle_query_tool.py"
if errorlevel 1 exit /b 1

echo Building CteMetadataExporter.exe...
"%PYTHON_EXE%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name CteMetadataExporter ^
    --distpath "%OUT_DIR%" ^
    --workpath "%WORK_DIR%" ^
    --specpath "%WORK_DIR%" ^
    --hidden-import config_loader ^
    --hidden-import oracle_client ^
    --hidden-import sql_reader ^
    --hidden-import excel_exporter ^
    --hidden-import chardet ^
    --hidden-import oracledb ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    "%ROOT%\export_cte_metadata.py"
if errorlevel 1 exit /b 1

copy /y "%ROOT%\config.example.json" "%OUT_DIR%\config.example.json" >nul

echo.
echo Build complete:
echo   %OUT_DIR%\OracleQueryTool.exe
echo   %OUT_DIR%\CteMetadataExporter.exe
echo.
echo To use on another Windows machine, copy the whole folder:
echo   %OUT_DIR%
echo Then copy config.example.json to config.json and edit Oracle settings.

endlocal
