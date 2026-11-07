
$ErrorActionPreference = "Continue"
$context = Get-Content -Raw -Encoding UTF8 "C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\s1f1_cf_parity_work\cf_batch_wps_context.json" | ConvertFrom-Json
$app = New-Object -ComObject KET.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$out = @{}
try {
    foreach ($entry in $context.files) {
        $fileResult = @{ cfSheets = @(); noCf = $false; error = "" }
        try {
            $wb = $app.Workbooks.Open($entry.path)
            try {
                if ($entry.sheets.Count -eq 0) {
                    $fileResult.noCf = $true
                } else {
                    $cfSheets = @()
                    $wb.Activate()
                    $app.ScreenUpdating = $true
                    foreach ($sheetEntry in $entry.sheets) {
                        $sheet = $wb.Worksheets.Item($sheetEntry.name)
                        $sheet.Activate()
                        try { $app.CalculateFullRebuild() } catch { try { $app.Calculate() } catch {} }
                        $triggered = @()
                        $scanned = 0
                        foreach ($condition in $sheet.UsedRange.FormatConditions) {
                            $appliesTo = $condition.AppliesTo
                            foreach ($area in $appliesTo.Areas) {
                                foreach ($cell in $area.Cells) {
                                    $scanned++
                                    $addr = $cell.Address($false,$false)
                                    $display = $null
                                    try { $display = $cell.DisplayFormat.Interior.Color } catch { }
                                    if ($null -eq $display) { continue }
                                    $base = 16777215
                                    try { $base = $cell.Interior.Color } catch {}
                                    if ($display -ne $base) {
                                        $triggered += ("$addr=$display")
                                    }
                                }
                            }
                        }
                        $cfSheets += @{ name = $sheetEntry.name; scanned = $scanned; triggered = $triggered }
                    }
                    $fileResult.cfSheets = $cfSheets
                    $app.ScreenUpdating = $false
                }
            } finally {
                $wb.Close($false)
            }
        } catch {
            $fileResult.error = $_.Exception.Message
        }
        $out[$entry.path] = $fileResult
    }
    $json = $out | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "BATCH-OK"
} finally {
    $app.Quit()
}
