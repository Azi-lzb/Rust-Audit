@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
if not exist "%DIST_DIR%\AuditTools.exe" (
    echo Missing %DIST_DIR%\AuditTools.exe. Run build-release.bat explicitly first.
    exit /b 2
)
if not exist "%PROJECT_ROOT%audit-core\config-real" (
    echo Missing local real configuration directory: %PROJECT_ROOT%audit-core\config-real
    echo Required files are checked for the selected edition by package-windows.ps1.
    exit /b 4
)
if not exist "%PROJECT_ROOT%package-windows.ps1" (
    echo Missing package-windows.ps1.
    exit /b 5
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_ROOT%package-windows.ps1" %*
exit /b %ERRORLEVEL%
