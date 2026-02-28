@echo off
chcp 65001 >nul
echo ========================================
echo   工厂订单转换工具 - 打包脚本
echo ========================================
echo.

:: 从 version.py 读取版本号
for /f "tokens=2 delims==" %%a in ('findstr "VERSION" version.py ^| findstr /v "APP_NAME"') do (
    set RAW_VER=%%a
)
:: 去掉引号和空格
set VER=%RAW_VER: =%
set VER=%VER:"=%
echo 当前版本: %VER%
echo.

echo [1/3] 安装依赖...
pip install -r requirements.txt pyinstaller

echo.
echo [2/3] 开始打包（单文件夹模式）...
pyinstaller --onedir --windowed --name "订单转换工具" ^
    --add-data "mapping_table.xlsx;." ^
    --clean ^
    main.py

echo.
echo [3/3] 整理输出...

:: 在dist目录创建带版本号的发布包
set RELEASE_DIR=dist\订单转换工具_v%VER%
if exist "%RELEASE_DIR%" rmdir /s /q "%RELEASE_DIR%"
rename "dist\订单转换工具" "订单转换工具_v%VER%"

echo.
echo ========================================
echo   打包完成！
echo   输出目录: %RELEASE_DIR%\
echo   版本: v%VER%
echo ========================================
echo.
echo 部署方式:
echo   将 %RELEASE_DIR%\ 整个文件夹复制到目标电脑即可使用
echo   mapping_table.xlsx 已在文件夹内
echo.
pause
