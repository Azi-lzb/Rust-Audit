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

echo Building AuditTools Release binaries from this source into %CARGO_TARGET_DIR%.
cargo build --release --offline --locked -p audit-core --bin AuditTools
if errorlevel 1 goto build_failed

if not exist "%PROJECT_ROOT%dist" mkdir "%PROJECT_ROOT%dist"
copy /y "%CARGO_TARGET_DIR%\release\AuditTools.exe" "%PROJECT_ROOT%dist\AuditTools.exe" >nul
if errorlevel 1 goto copy_failed
set "OPENPYXL_WORKER_DIR=%PROJECT_ROOT%dist\libexec\audittools\openpyxl\worker"
if not exist "%OPENPYXL_WORKER_DIR%" mkdir "%OPENPYXL_WORKER_DIR%"
copy /y "%PROJECT_ROOT%crates\excel-editor\python\openpyxl_worker.py" "%OPENPYXL_WORKER_DIR%\openpyxl_worker.py" >nul
if errorlevel 1 goto copy_failed
if exist "%CARGO_TARGET_DIR%\release\VERSION_MANIFEST.json" (
    copy /y "%CARGO_TARGET_DIR%\release\VERSION_MANIFEST.json" "%PROJECT_ROOT%dist\VERSION_MANIFEST.json" >nul
) else (
    copy /y "%PROJECT_ROOT%VERSION_MANIFEST.json" "%PROJECT_ROOT%dist\VERSION_MANIFEST.json" >nul
)
echo AuditTools Release EXE copied to %PROJECT_ROOT%dist.
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
