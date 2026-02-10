@echo off
REM DOCX Batch Updater - Build Script
REM 打包为单文件 exe

echo ====================================
echo DOCX 批量更新器 - 开始打包
echo ====================================
echo.

REM 检查 Python
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python 未安装或不在 PATH 中
    pause
    exit /b 1
)

REM 安装依赖
echo.
echo [1/3] 安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: 依赖安装失败
    pause
    exit /b 1
)

REM 清理旧构建
echo.
echo [2/3] 清理旧构建文件...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "*.spec" del /q "*.spec"

REM 执行打包
echo.
echo [3/3] 执行 PyInstaller 打包...
pyinstaller --name="DOCX Batch Updater" --windowed --onefile ^
    --add-data="src;src" ^
    --hidden-import=PyQt5.sip ^
    --hidden-import=docx ^
    --hidden-import=docx.opc.constants ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.text.paragraph ^
    --hidden-import=docx.oxml.table ^
    src/main.py

if %errorlevel% neq 0 (
    echo ERROR: 打包失败
    pause
    exit /b 1
)

echo.
echo ====================================
echo 打包完成！
echo ====================================
echo.
echo 输出文件: dist\DOCX Batch Updater.exe
echo 大小: 约 39 MB
echo.
echo 双击 exe 文件即可运行（无需安装 Python）
echo ====================================
echo.
pause
