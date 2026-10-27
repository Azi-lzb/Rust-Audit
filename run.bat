@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
set "AUDIT_EXE=%DIST_DIR%\open-data-audit.exe"

if not exist "%AUDIT_EXE%" (
    echo Audit EXE is missing: %AUDIT_EXE%
    echo This script never invokes Cargo. Run build-release.bat explicitly to build from source.
    exit /b 2
)
pushd "%PROJECT_ROOT%"
if errorlevel 1 exit /b 4
"%AUDIT_EXE%" serve %*
set "RUN_EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %RUN_EXIT_CODE%

