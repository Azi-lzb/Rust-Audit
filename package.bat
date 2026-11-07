@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
if not exist "%DIST_DIR%\open-data-audit.exe" (
    echo Missing %DIST_DIR%\open-data-audit.exe. Run build-release.bat explicitly first.
    exit /b 2
)
if not exist "%PROJECT_ROOT%audit-core\config-real" (
    echo Missing local real configuration directory: %PROJECT_ROOT%audit-core\config-real
    echo Copy the six S1, S2, and S3 configuration workbooks there before packaging.
    exit /b 4
)
if not exist "%PROJECT_ROOT%package-windows.ps1" (
    echo Missing package-windows.ps1.
    exit /b 5
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_ROOT%package-windows.ps1"
exit /b %ERRORLEVEL%
