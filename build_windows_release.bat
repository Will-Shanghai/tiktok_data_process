@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem Build from the repository root containing this script.
cd /d "%~dp0"

for /f "usebackq tokens=* delims=" %%V in (`python -c "from pathlib import Path; print(Path('VERSION').read_text(encoding='utf-8-sig').strip())"`) do if not defined TOOL_VERSION set "TOOL_VERSION=%%V"
if not defined TOOL_VERSION (
  echo ERROR: VERSION is missing or empty.
  exit /b 1
)

set "EXE_BASE=TikTokShopDataTool_v%TOOL_VERSION%"
set "DIST_ROOT=dist\%EXE_BASE%"
set "RELEASE_ROOT=release\%EXE_BASE%"
set "ZIP_PATH=release\%EXE_BASE%_windows.zip"
set "USAGE_GUIDE="

rem Avoid a hard-coded Chinese file name because CMD code pages differ by Windows setup.
for %%F in (TikTokShopDataTool_*.txt) do if not defined USAGE_GUIDE set "USAGE_GUIDE=%%~fF"
if not defined USAGE_GUIDE (
  echo ERROR: The TikTokShopDataTool usage guide (*.txt) is missing.
  exit /b 1
)

for %%F in (
  "main.py"
  "sum_daily_order\config\app_config.xlsx"
  "sum_daily_conversion\config\config.json"
  "sum_daily_conversion\config\product_sheet_mapping.csv"
) do (
  if not exist %%F (
    echo ERROR: Required file is missing: %%~F
    exit /b 1
  )
)

if exist "%RELEASE_ROOT%" (
  echo ERROR: Release directory already exists: %RELEASE_ROOT%
  echo Move it elsewhere or increase VERSION before building. It is not overwritten.
  exit /b 1
)

echo Building %EXE_BASE% ...
python -m PyInstaller --clean --noconfirm --onedir --name "%EXE_BASE%" ^
  --hidden-import sum_daily_order.config.cal_cost_return_jp_daily ^
  --hidden-import sum_daily_order.config.cal_cost_return_vn_daily ^
  --hidden-import sum_daily_order.config.cal_cost_return_mx_daily ^
  --hidden-import sum_daily_conversion.config.cal_product_daily_conversion ^
  main.py
if errorlevel 1 goto :failed

mkdir "%RELEASE_ROOT%\config" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\config\cache" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\local" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\cross_border" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\direct_old" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_JP\direct_new" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_VN\local" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_VN\cross_border" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\local" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\direct_old" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\data\data_MX\direct_new" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_order\result" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_conversion\config" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_conversion\data" || goto :failed
mkdir "%RELEASE_ROOT%\sum_daily_conversion\result\logs" || goto :failed

xcopy "%DIST_ROOT%\*" "%RELEASE_ROOT%\" /E /I /Y /Q >nul || goto :failed
copy /Y "VERSION" "%RELEASE_ROOT%\VERSION" >nul || goto :failed
copy /Y "%USAGE_GUIDE%" "%RELEASE_ROOT%\" >nul || goto :failed
copy /Y "sum_daily_order\config\app_config.xlsx" "%RELEASE_ROOT%\sum_daily_order\config\" >nul || goto :failed
copy /Y "sum_daily_order\config\.env.example" "%RELEASE_ROOT%\config\.env.example" >nul || goto :failed
copy /Y "sum_daily_conversion\config\config.json" "%RELEASE_ROOT%\sum_daily_conversion\config\" >nul || goto :failed
copy /Y "sum_daily_conversion\config\product_sheet_mapping.csv" "%RELEASE_ROOT%\sum_daily_conversion\config\" >nul || goto :failed
if exist "sum_daily_order\config\cache\*" xcopy "sum_daily_order\config\cache\*" "%RELEASE_ROOT%\sum_daily_order\config\cache\" /E /I /Y /Q >nul

if exist "%ZIP_PATH%" del /q "%ZIP_PATH%"
powershell -NoProfile -Command "Compress-Archive -Path '%RELEASE_ROOT%' -DestinationPath '%ZIP_PATH%' -Force"
if errorlevel 1 goto :failed

echo.
echo SUCCESS: %ZIP_PATH%
echo Distribute the ZIP file. Do not distribute dist\%EXE_BASE% by itself.
exit /b 0

:failed
echo.
echo ERROR: Build failed. No release ZIP was created.
exit /b 1
