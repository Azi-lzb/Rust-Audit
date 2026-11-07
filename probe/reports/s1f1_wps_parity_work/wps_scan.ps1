
$ErrorActionPreference = "Stop"
$context = Get-Content -Raw -Encoding UTF8 "C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\s1f1_wps_parity_work\wps_context.json" | ConvertFrom-Json

# 启动 WPS 表格（单实例，非只读）
$app = New-Object -ComObject KET.Application
$app.Visible = $false
$app.DisplayAlerts = $false

$out = @{}
try {
    foreach ($entry in $context.files) {
        $wb = $app.Workbooks.Open($entry.path)
        try {
            # 关键：激活工作簿 + 开启屏幕更新，让 WPS 渲染条件格式
            $app.ScreenUpdating = $true
            $wb.Activate()

            $cfSheets = @()
            foreach ($sheetEntry in $entry.sheets) {
                $sheet = $wb.Worksheets.Item($sheetEntry.name)

                # 逐表激活 + 重算（旧 Python Audit 的 WPS 刷新配方）
                $sheet.Activate()
                $app.CalculateFullRebuild()

                $triggered = @()
                $scanned = 0
                foreach ($condition in $sheet.UsedRange.FormatConditions) {
                    $appliesTo = $condition.AppliesTo
                    foreach ($area in $appliesTo.Areas) {
                        foreach ($cell in $area.Cells) {
                            $scanned++
                            $addr = $cell.Address($false,$false)
                            try {
                                $display = $cell.DisplayFormat.Interior.Color
                                $base = $cell.Interior.Color
                                if ($display -ne $base) {
                                    $triggered += ("$addr=$display")
                                }
                            } catch { }
                        }
                    }
                }
                $cfSheets += @{
                    name = $sheetEntry.name
                    scanned = $scanned
                    triggered = $triggered
                }
            }
            $out[$entry.path] = @{ cfSheets = $cfSheets; error = "" }
        } catch {
            $out[$entry.path] = @{ error = $_.Exception.Message; cfSheets = @() }
        } finally {
            $wb.Close($false)
        }
    }
    $json = $out | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "WPS-BATCH-OK"
} finally {
    $app.Quit()
    $app = $null
    [System.GC]::Collect()
}
