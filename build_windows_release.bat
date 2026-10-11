@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0"

set /p TOOL_VERSION=<VERSION
if not defined TOOL_VERSION goto missing_version

set EXE_BASE=TikTokShopDataTool_v%TOOL_VERSION%
set DIST_ROOT=dist\%EXE_BASE%
set RELEASE_ROOT=release\%EXE_BASE%
set ZIP_PATH=release\%EXE_BASE%_windows.zip

if not exist main.py goto missing_main
if not exist sum_daily_order\config\app_config.xlsx goto missing_order_config
if not exist sum_daily_conversion\config\config.json goto missing_conversion_config
if not exist sum_daily_conversion\config\product_sheet_mapping.csv goto missing_mapping
if exist "%RELEASE_ROOT%" goto release_exists

echo Building %EXE_BASE% ...
python -m PyInstaller --clean --noconfirm --onedir --name "%EXE_BASE%" --hidden-import sum_daily_order.config.cal_cost_return_jp_daily --hidden-import sum_daily_order.config.cal_cost_return_vn_daily --hidden-import sum_daily_order.config.cal_cost_return_mx_daily --hidden-import sum_daily_conversion.config.cal_product_daily_conversion main.py
if errorlevel 1 goto failed

mkdir "%RELEASE_ROOT%\config"
mkdir "%RELEASE_ROOT%\sum_daily_order\config\cache"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\local"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\cross_border"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\direct_old"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\direct_new"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_VN\local"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_VN\cross_border"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\local"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\direct_old"
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\direct_new"
mkdir "%RELEASE_ROOT%\sum_daily_order\result"
mkdir "%RELEASE_ROOT%\sum_daily_conversion\config"
mkdir "%RELEASE_ROOT%\sum_daily_conversion\data"
mkdir "%RELEASE_ROOT%\sum_daily_conversion\result\logs"

xcopy "%DIST_ROOT%\*" "%RELEASE_ROOT%\" /E /I /Y /Q >nul
if errorlevel 1 goto failed
copy /Y VERSION "%RELEASE_ROOT%\VERSION" >nul
copy /Y sum_daily_order\config\app_config.xlsx "%RELEASE_ROOT%\sum_daily_order\config\" >nul
copy /Y sum_daily_order\config\.env.example "%RELEASE_ROOT%\config\.env.example" >nul
copy /Y sum_daily_conversion\config\config.json "%RELEASE_ROOT%\sum_daily_conversion\config\" >nul
copy /Y sum_daily_conversion\config\product_sheet_mapping.csv "%RELEASE_ROOT%\sum_daily_conversion\config\" >nul
if exist sum_daily_order\config\cache\* xcopy sum_daily_order\config\cache\* "%RELEASE_ROOT%\sum_daily_order\config\cache\" /E /I /Y /Q >nul

powershell -NoProfile -Command "Compress-Archive -Path '%RELEASE_ROOT%' -DestinationPath '%ZIP_PATH%' -Force"
if errorlevel 1 goto failed

echo.
echo SUCCESS: %ZIP_PATH%
exit /b 0

:missing_version
echo ERROR: VERSION is missing or empty.
exit /b 1
:missing_main
echo ERROR: main.py is missing.
exit /b 1
:missing_order_config
echo ERROR: sum_daily_order config is missing.
exit /b 1
:missing_conversion_config
echo ERROR: sum_daily_conversion config.json is missing.
exit /b 1
:missing_mapping
echo ERROR: product_sheet_mapping.csv is missing.
exit /b 1
:release_exists
echo ERROR: Release directory already exists: %RELEASE_ROOT%
echo Update VERSION or rename the existing release directory, then run again.
exit /b 1
:failed
echo ERROR: Build failed. Check the messages above.
exit /b 1
