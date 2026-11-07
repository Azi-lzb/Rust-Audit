
$ErrorActionPreference = "Stop"
$ctx = Get-Content -Raw -Encoding UTF8 "C:\Users\AZI\Desktop\VibeCode\Rust-Audit\probe\reports\poison_template_work\poison_copies_context.json" | ConvertFrom-Json
$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$result = @{}
try {
    foreach ($path in $ctx.files) {
        try {
            $wb = $app.Workbooks.Open($path, 0, $true)
            $result[$path] = "OK sheets=" + $wb.Worksheets.Count
            $wb.Close($false)
        } catch {
            $msg = $_.Exception.Message
            $result[$path] = "FAIL " + $msg.Substring(0, [Math]::Min(120, $msg.Length))
        }
    }
    $json = $result | ConvertTo-Json -Depth 3
    [System.IO.File]::WriteAllText($ctx.out, $json, (New-Object System.Text.UTF8Encoding($true)))
    Write-Output "OPEN-BATCH-OK"
} finally {
    $app.Quit()
}
