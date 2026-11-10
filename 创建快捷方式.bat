@echo off
chcp 65001 >nul
setlocal

rem 在开始菜单（和可选的桌面）创建「审核工具」快捷方式。
rem 免安装布局：本脚本只写当前用户的快捷方式，不需要管理员权限。

set "APP_DIR=%~dp0"
set "EXE=%APP_DIR%AuditTools.exe"
if not exist "%EXE%" (
    echo [错误] 未找到 AuditTools.exe，请把本脚本与 EXE 放在同一文件夹。
    pause
    exit /b 2
)

set "START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
set "DESKTOP=%USERPROFILE%\Desktop"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$exe = '%EXE%';" ^
  "$dir = '%APP_DIR%';" ^
  "$targets = @((Join-Path '%START_MENU%' '审核工具.lnk'));" ^
  "if (Test-Path '%DESKTOP%') { $targets += (Join-Path '%DESKTOP%' '审核工具.lnk') };" ^
  "$shell = New-Object -ComObject WScript.Shell;" ^
  "foreach ($lnk in $targets) {" ^
  "  $sc = $shell.CreateShortcut($lnk);" ^
  "  $sc.TargetPath = $exe;" ^
  "  $sc.WorkingDirectory = $dir;" ^
  "  $sc.IconLocation = $exe + ',0';" ^
  "  $sc.Description = '审核工具：逐笔统计 / 报表采集 / 大集中统计审核';" ^
  "  $sc.WindowStyle = 7;" ^
  "  $sc.Save();" ^
  "  Write-Host ('已创建：' + $lnk)" ^
  "}"

if errorlevel 1 (
    echo [错误] 创建快捷方式失败。
    pause
    exit /b 3
)

echo.
echo 完成：已添加到「开始菜单」，并在桌面创建快捷方式（如存在桌面目录）。
echo 之后可从开始菜单搜索「审核工具」启动，无需再次运行本脚本。
echo.
pause
exit /b 0
