$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$dist = Join-Path $root 'dist'
$configSource = Join-Path $root 'audit-core\config-real'
$configFiles = @(
    '1.逐笔统计系统_配置.xlsx',
    '2.报表采集系统_配置.xlsx',
    '3.0大集中通用配置.xlsx',
    '3.1大集中执行比较_配置.xlsx',
    '3.2大集中本期数值核对_配置.xlsx',
    '3.3大集中指标比较拆分_配置.xlsx'
)

foreach ($name in $configFiles) {
    $source = Join-Path $configSource $name
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Required S1/S2/S3 configuration is missing: $source"
    }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$archive = Join-Path $dist "open-data-audit-windows-x64-$stamp.zip"
if (Test-Path -LiteralPath $archive) {
    throw "Package already exists: $archive"
}

$stage = Join-Path ([IO.Path]::GetTempPath()) ('audit-package-' + [Guid]::NewGuid().ToString('N'))
try {
    $stageDist = Join-Path $stage 'dist'
    $stageConfig = Join-Path $stageDist 'config-real'
    New-Item -ItemType Directory -Path $stageConfig -Force | Out-Null

    Copy-Item -LiteralPath (Join-Path $dist 'open-data-audit.exe') -Destination $stageDist
    $workerDirectory = Join-Path $stageDist 'libexec\open-data-audit\openpyxl\worker'
    New-Item -ItemType Directory -Path $workerDirectory -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $root 'crates\excel-editor\python\openpyxl_worker.py') -Destination (Join-Path $workerDirectory 'openpyxl_worker.py')
    Copy-Item -LiteralPath (Join-Path $root 'run.bat') -Destination $stage
    Copy-Item -LiteralPath (Join-Path $root 'README.md') -Destination $stage
    foreach ($name in $configFiles) {
        Copy-Item -LiteralPath (Join-Path $configSource $name) -Destination $stageConfig
    }

    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $archive
    Write-Host "Package created: $archive"
    Write-Host "Included $($configFiles.Count) real S1/S2/S3 configuration workbooks."
}
finally {
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
}
