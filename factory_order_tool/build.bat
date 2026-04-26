@echo off
chcp 65001 >nul
echo ========================================
echo   工厂订单转换工具 - 打包脚本
echo ========================================
echo.

:: ====================================================
:: 读取版本号 + 构建日期 (version.py 是唯一来源)
:: ====================================================
for /f "tokens=2 delims==" %%a in ('findstr "^VERSION" version.py') do set RAW_VER=%%a
for /f "tokens=2 delims==" %%a in ('findstr "^BUILD_DATE" version.py') do set RAW_DATE=%%a
set VER=%RAW_VER: =%
set VER=%VER:"=%
set BUILD_DATE=%RAW_DATE: =%
set BUILD_DATE=%BUILD_DATE:"=%
echo 当前版本: v%VER%   (构建日期: %BUILD_DATE%)
echo.

:: 指定 Python 3.10（PyInstaller 与 Python 3.14 存在兼容性问题）
set PY310=C:\Users\LEGION\AppData\Local\Programs\Python\Python310\python.exe

:: ====================================================
:: [1/5] 虚拟环境
:: ====================================================
if not exist ".venv\Scripts\python.exe" (
    echo [1/5] 创建干净的虚拟环境（Python 3.10）...
    "%PY310%" -m venv .venv
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\pip.exe install -r requirements.txt pyinstaller
) else (
    echo [1/5] 虚拟环境已存在，跳过创建
)

:: ====================================================
:: [2/5] 清理旧构建
:: ====================================================
echo.
echo [2/5] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "订单转换工具.spec" del "订单转换工具.spec"

:: ====================================================
:: [3/5] PyInstaller 打包
:: ====================================================
echo.
echo [3/5] 开始打包（单文件夹模式）...
.venv\Scripts\python.exe -m PyInstaller --onedir --windowed --name "订单转换工具" ^
    --clean ^
    main.py

if errorlevel 1 (
    echo.
    echo [X] PyInstaller 打包失败，请检查上方错误信息。
    pause
    exit /b 1
)

:: ====================================================
:: [4/5] 整理 完整包（首次部署 / 新客户 / 应急重置 用）
:: ====================================================
echo.
echo [4/5] 整理 完整包（含 mapping_table.xlsx）...
set FULL_DIR=dist\订单转换工具_v%VER%_完整包
rename "dist\订单转换工具" "订单转换工具_v%VER%_完整包"

if exist "mapping_table.xlsx" (
    copy /Y "mapping_table.xlsx" "%FULL_DIR%\mapping_table.xlsx" >nul
    echo   [OK] 已复制 mapping_table.xlsx 到 完整包
) else (
    echo   [!] 仓库无 mapping_table.xlsx，完整包不带映射表
    echo       此情况下完整包仅适用于新客户「自带映射表首次部署」
)

if exist "CHANGELOG.md" (
    copy /Y "CHANGELOG.md" "%FULL_DIR%\CHANGELOG.md" >nul
    echo   [OK] 已复制 CHANGELOG.md 到 完整包
)

:: ====================================================
:: [5/5] 派生 升级包（老客户升级用，不含 mapping_table.xlsx）
:: ====================================================
echo.
echo [5/5] 派生 升级包（不含 mapping_table.xlsx）...
set UP_DIR=dist\订单转换工具_v%VER%_升级包
mkdir "%UP_DIR%"
copy /Y "%FULL_DIR%\订单转换工具.exe" "%UP_DIR%\" >nul
xcopy /E /I /Q /Y "%FULL_DIR%\_internal" "%UP_DIR%\_internal\" >nul
echo   [OK] 已复制 订单转换工具.exe + _internal\ 到 升级包

:: 渲染 部署说明.txt （用 PowerShell 替换占位符）
if exist "release\部署说明_template.txt" (
    powershell -NoProfile -Command "(Get-Content 'release\部署说明_template.txt' -Raw -Encoding UTF8) -replace '{VERSION}','%VER%' -replace '{BUILD_DATE}','%BUILD_DATE%' | Set-Content -Path '%UP_DIR%\部署说明.txt' -Encoding UTF8"
    echo   [OK] 已生成 部署说明.txt（已渲染版本号 %VER% / 构建日期 %BUILD_DATE%）
) else (
    echo   [!] 未找到 release\部署说明_template.txt，升级包将不含部署说明
)

:: 复制 CHANGELOG 让客户能看到完整变更历史
if exist "CHANGELOG.md" (
    copy /Y "CHANGELOG.md" "%UP_DIR%\CHANGELOG.md" >nul
    echo   [OK] 已复制 CHANGELOG.md 到 升级包
)

:: 复制升级助手 (bat 引导 + ps1 主逻辑 + 用户数据白名单)
if exist "release\升级.bat" (
    copy /Y "release\升级.bat" "%UP_DIR%\升级.bat" >nul
    echo   [OK] 已复制 升级.bat 到 升级包
)
if exist "release\upgrade.ps1" (
    copy /Y "release\upgrade.ps1" "%UP_DIR%\upgrade.ps1" >nul
    echo   [OK] 已复制 upgrade.ps1 到 升级包
)
if exist "release\user_data_files.txt" (
    copy /Y "release\user_data_files.txt" "%UP_DIR%\user_data_files.txt" >nul
    echo   [OK] 已复制 user_data_files.txt（用户数据白名单）到 升级包
)

:: ====================================================
:: 完成
:: ====================================================
echo.
echo ========================================
echo   打包完成! v%VER% / %BUILD_DATE%
echo ========================================
echo.
echo   完整包: %FULL_DIR%\
echo           （含 mapping_table.xlsx，首次部署 / 新客户 / 应急重置 用）
echo.
echo   升级包: %UP_DIR%\
echo           （给已在使用的老客户用，自带升级助手）
echo           客户解压后双击「升级.bat」，弹窗选择旧部署目录，
echo           脚本自动备份旧版 + 迁移 mapping_table / settings.json
echo           + 部署新版 + 启动验证。
echo.
pause
