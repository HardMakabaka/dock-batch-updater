@echo off
REM Build script for DOCX Batch Updater using PyInstaller

echo ====================================
echo DOCX Batch Updater - Build Script
echo ====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    exit /b 1
)

echo Python found:
python --version
echo.

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    exit /b 1
)
echo Dependencies installed successfully
echo.

REM Clean previous build
echo Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"
echo Previous build cleaned
echo.

REM Build with PyInstaller
echo Building executable with PyInstaller...
pyinstaller --name="DOCX Batch Updater" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data="src;src" ^
    --hidden-import=PyQt5.sip ^
    --hidden-import=docx ^
    --hidden-import=docx.opc.constants ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.text.paragraph ^
    --hidden-import=docx.oxml.table ^
    src/main.py

if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed
    exit /b 1
)
echo Build completed successfully
echo.

REM Show output location
echo ====================================
echo Build completed!
echo ====================================
echo Executable location: dist\DOCX Batch Updater.exe
echo.
echo You can now distribute the executable file.
echo Note: The executable includes all dependencies and
echo       can be run on Windows without Python installed.
echo ====================================

pause
