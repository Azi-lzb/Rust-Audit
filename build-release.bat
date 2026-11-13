@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "CARGO_TARGET_DIR=F:\RustBuild\audit"
set "AUDIT_BUILD_EDITION=%~1"
if "%AUDIT_BUILD_EDITION%"=="" set "AUDIT_BUILD_EDITION=full"
set "AUDIT_FEATURE_ARGS="
if /I "%AUDIT_BUILD_EDITION%"=="s1s2" (
    set "AUDIT_FEATURE_ARGS=--features windows-s1s2"
) else if /I not "%AUDIT_BUILD_EDITION%"=="full" (
    echo Usage: build-release.bat [full^|s1s2]
    exit /b 2
)
where cargo >nul 2>nul
if errorlevel 1 (
    echo cargo was not found. No build was started.
    exit /b 9009
)

pushd "%PROJECT_ROOT%"
if errorlevel 1 exit /b 3

echo Building AuditTools Release binaries from this source into %CARGO_TARGET_DIR%.
cargo build --release --offline --locked -p audit-core --bin AuditTools %AUDIT_FEATURE_ARGS%
if errorlevel 1 goto build_failed

if not exist "%PROJECT_ROOT%dist" mkdir "%PROJECT_ROOT%dist"
copy /y "%CARGO_TARGET_DIR%\release\AuditTools.exe" "%PROJECT_ROOT%dist\AuditTools.exe" >nul
if errorlevel 1 goto copy_failed
set "OPENPYXL_WORKER_DIR=%PROJECT_ROOT%dist\libexec\audittools\openpyxl\worker"
if not exist "%OPENPYXL_WORKER_DIR%" mkdir "%OPENPYXL_WORKER_DIR%"
copy /y "%PROJECT_ROOT%crates\excel-editor\python\openpyxl_worker.py" "%OPENPYXL_WORKER_DIR%\openpyxl_worker.py" >nul
if errorlevel 1 goto copy_failed
powershell.exe -NoProfile -Command "$ErrorActionPreference='Stop'; $exe=Join-Path $env:PROJECT_ROOT 'dist\AuditTools.exe'; $stream=[IO.File]::OpenRead($exe); $hasher=[Security.Cryptography.SHA256]::Create(); try { $digest=[BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-','') } finally { $stream.Dispose(); $hasher.Dispose() }; @{schema_version=1; edition=$env:AUDIT_BUILD_EDITION.ToLowerInvariant(); sha256=$digest} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $env:PROJECT_ROOT 'dist\AuditTools.build.json') -Encoding ascii"
if errorlevel 1 goto copy_failed
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
