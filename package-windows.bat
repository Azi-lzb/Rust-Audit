@echo off
rem 打包 Windows 发行目录：根目录放 exe，配置放 config/默认配置，不带 bat。
rem 只使用 dist 里已存在的构建产物，不触发任何 Cargo 构建。
setlocal EnableDelayedExpansion

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
set "PACKAGE_NAME=OpenDataAudit"
set "PACKAGE_DIR=%DIST_DIR%\%PACKAGE_NAME%"

if not exist "%DIST_DIR%\open-data-audit.exe" (
    echo open-data-audit.exe is missing in %DIST_DIR%.
    echo Run build-release.bat explicitly to build it first. No build was started.
    exit /b 2
)
if not exist "%PROJECT_ROOT%audit-core\config\默认配置" (
    echo Default config directory is missing: audit-core\config\默认配置
    exit /b 3
)

if exist "%PACKAGE_DIR%" rmdir /s /q "%PACKAGE_DIR%"
mkdir "%PACKAGE_DIR%"
if errorlevel 1 exit /b 4

copy /y "%DIST_DIR%\open-data-audit.exe" "%PACKAGE_DIR%\open-data-audit.exe" >nul
if errorlevel 1 exit /b 5

mkdir "%PACKAGE_DIR%\config\默认配置"
xcopy /e /i /y /q "%PROJECT_ROOT%audit-core\config\默认配置\*" "%PACKAGE_DIR%\config\默认配置\" >nul
if errorlevel 1 exit /b 6

if exist "%PROJECT_ROOT%README_RELEASE.md" (
    copy /y "%PROJECT_ROOT%README_RELEASE.md" "%PACKAGE_DIR%\README_RELEASE.md" >nul
) else if exist "%DIST_DIR%\README_RELEASE.md" (
    copy /y "%DIST_DIR%\README_RELEASE.md" "%PACKAGE_DIR%\README_RELEASE.md" >nul
)

rem 内置 openpyxl worker 的隔离 Python 运行时（openpyxl 备选引擎）。
rem 用系统 Python（须已装 openpyxl）复制精简运行时；失败不阻断打包，
rem 但会明确提示 openpyxl 备选在目标机器不可用。
if exist "%PROJECT_ROOT%bundle_openpyxl_runtime.py" (
  python "%PROJECT_ROOT%bundle_openpyxl_runtime.py" --dest "%PACKAGE_DIR%\libexec\open-data-audit\openpyxl/runtime"
  if errorlevel 1 (
    echo WARNING: bundled python runtime failed; openpyxl fallback will need PATH python.
  ) else (
    if not exist "%PACKAGE_DIR%\libexec/open-data-audit/openpyxl/worker" mkdir "%PACKAGE_DIR%\libexec/open-data-audit/openpyxl/worker"
    if exist "%PROJECT_ROOT%crates/excel-editor/python/openpyxl_worker.py" copy /y "%PROJECT_ROOT%crates/excel-editor/python/openpyxl_worker.py" "%PACKAGE_DIR%\libexec/open-data-audit/openpyxl/worker\openpyxl_worker.py" >nul
    echo Bundled python+openpyxl runtime staged.
  )
)

echo Packaged to %PACKAGE_DIR%
echo Layout: open-data-audit.exe (root) + config\默认配置\
dir /b "%PACKAGE_DIR%"
popd 2>nul
exit /b 0
