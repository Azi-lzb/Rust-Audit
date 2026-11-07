
$ErrorActionPreference = "Stop"
$context = Get-Content -Raw -Encoding UTF8 "C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\cf_scan_wps_context.json" | ConvertFrom-Json
$app = $null
$used = ""
foreach ($name in $context.probes) {
    try { $app = New-Object -ComObject $name; $used = $name; break } catch {}
}
if (-not $app) { Write-Output "NO_ENGINE"; exit 1 }
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Open($context.workbook, 0, $true)
    try {
        $app.CalculateFullRebuild()
        $result = @{}
        foreach ($sheetName in $context.sheetNames) {
            $sheet = $wb.Worksheets.Item($sheetName)
            $cfFormats = $sheet.Cells.FormatConditions
            $count = $cfFormats.Count
            $triggered = @()
            $scanned = 0
            for ($i = 1; $i -le $count; $i++) {
                $fmt = $cfFormats.Item($i)
                $appliesTo = $null
                try { $appliesTo = $fmt.Parent.Range($fmt.AppliesTo.Address($true,$true)) } catch { continue }
                foreach ($cell in $appliesTo.Cells) {
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
            $result[$sheetName] = @{ scanned = $scanned; triggeredCount = $triggered.Count; triggered = $triggered }
        }
        $json = $result | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($context.scanOut, $json, (New-Object System.Text.UTF8Encoding($true)))
        Write-Output "COM-OK engine=$used"
    } finally {
        $wb.Close($false)
        $app.Quit()
    }
} catch {
    Write-Output ("COM-FAIL: " + $_.Exception.Message)
    exit 1
}
