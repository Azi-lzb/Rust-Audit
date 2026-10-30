@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "CARGO_TARGET_DIR=F:\RustBuild\audit"
where cargo >nul 2>nul
if errorlevel 1 (
    echo cargo was not found. No build was started.
    exit /b 9009
)

pushd "%PROJECT_ROOT%"
if errorlevel 1 exit /b 3

echo Building Audit Release binaries from this source into %CARGO_TARGET_DIR%.
cargo build --release --offline --locked -p audit-core --bin open-data-audit
if errorlevel 1 goto build_failed

if not exist "%PROJECT_ROOT%dist" mkdir "%PROJECT_ROOT%dist"
copy /y "%CARGO_TARGET_DIR%\release\open-data-audit.exe" "%PROJECT_ROOT%dist\open-data-audit.exe" >nul
if errorlevel 1 goto copy_failed
echo Audit Release EXE copied to %PROJECT_ROOT%dist.
popd
exit /b 0

:build_failed
set "BUILD_EXIT_CODE=%ERRORLEVEL%"
popd
echo Release build failed with exit code %BUILD_EXIT_CODE%.
exit /b %BUILD_EXIT_CODE%

:copy_failed
set "COPY_EXIT_CODE=%ERRORLEVEL%"
popd
echo Build succeeded, but copying the EXE into dist failed with exit code %COPY_EXIT_CODE%.
exit /b %COPY_EXIT_CODE%

