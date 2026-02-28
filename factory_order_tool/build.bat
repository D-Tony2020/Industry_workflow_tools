@echo off
chcp 65001 >nul
echo ========================================
echo   工厂订单转换工具 - 打包脚本
echo ========================================
echo.

:: 从 version.py 读取版本号
for /f "tokens=2 delims==" %%a in ('findstr "^VERSION" version.py') do (
    set RAW_VER=%%a
)
set VER=%RAW_VER: =%
set VER=%VER:"=%
echo 当前版本: v%VER%
echo.

:: 检查虚拟环境
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] 创建干净的虚拟环境...
    python -m venv .venv
    .venv\Scripts\pip.exe install --upgrade pip
    .venv\Scripts\pip.exe install -r requirements.txt pyinstaller
) else (
    echo [1/4] 虚拟环境已存在，跳过创建
)

echo.
echo [2/4] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "订单转换工具.spec" del "订单转换工具.spec"

echo.
echo [3/4] 开始打包（单文件夹模式）...
.venv\Scripts\python.exe -m PyInstaller --onedir --windowed --name "订单转换工具" ^
    --add-data "mapping_table.xlsx;." ^
    --clean ^
    main.py

echo.
echo [4/4] 整理输出...
set RELEASE_DIR=dist\订单转换工具_v%VER%
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
rename "dist\订单转换工具" "订单转换工具_v%VER%"

echo.
echo ========================================
echo   打包完成!
echo   输出目录: %RELEASE_DIR%\
echo   版本: v%VER%
echo ========================================
echo.
echo 部署: 将 %RELEASE_DIR%\ 整个文件夹复制到目标电脑即可
echo.
pause
