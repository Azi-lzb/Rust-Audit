# AuditTools 审核工具 Windows 发布包（全量版 / S1S2 版）
#
# 用法：
#   powershell -File package-windows.ps1                  # 全量版 S1+S2+S3
#   powershell -File package-windows.ps1 -Edition s1s2    # S1+S2 版
#
# 布局（用户解压即可双击运行）：
#   <root>/
#     AuditTools.exe            双击即启动界面（无参数默认 serve）
#     config-real/*.xlsx        真实配置（S1S2 版只含 1./2.）
#     libexec/.../openpyxl_worker.py
#     使用说明.docx
#     [edition-s1s2]            仅 S1S2 版：界面只显示 S1/S2

param(
    [ValidateSet('full', 's1s2')]
    [string]$Edition = 'full'
)

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
$dist = Join-Path $root 'dist'
$configSource = Join-Path $root 'audit-core\config-real'
$docsDir = Join-Path $root 'docs'

$configFiles = @(
    '1.逐笔统计系统_配置.xlsx',
    '2.报表采集系统_配置.xlsx',
    '3.0大集中通用配置.xlsx',
    '3.1大集中执行比较_配置.xlsx',
    '3.2大集中本期数值核对_配置.xlsx',
    '3.3大集中指标比较拆分_配置.xlsx'
)
$s1s2ConfigFiles = @(
    '1.逐笔统计系统_配置.xlsx',
    '2.报表采集系统_配置.xlsx'
)

foreach ($name in $configFiles) {
    $source = Join-Path $configSource $name
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Required configuration is missing: $source"
    }
}

# GUI 默认绑定（settings.json 播种）依赖这两项随包分发：
# S1/S2 的「大集中系统数据」外部文件与 S1 模板文件目录。
$externalSource = Join-Path $configSource '外部文件-大集中相关数据.xlsx'
if (-not (Test-Path -LiteralPath $externalSource -PathType Leaf)) {
    throw "Required external data is missing: $externalSource"
}
$templateDirSource = Join-Path $configSource '模板文件_命名区域'
if (-not (Test-Path -LiteralPath $templateDirSource -PathType Container)) {
    throw "Required template directory is missing: $templateDirSource"
}

if ($Edition -eq 's1s2') {
    $includedConfigs = $s1s2ConfigFiles
    $guideName = '使用说明-S1S2版.docx'
    $editionTag = 's1s2'
} else {
    $includedConfigs = $configFiles
    $guideName = '使用说明-全量版.docx'
    $editionTag = 'full'
}

$guide = Join-Path $docsDir $guideName
if (-not (Test-Path -LiteralPath $guide -PathType Leaf)) {
    throw "使用说明缺失：$guide（先运行 probe/make_docs.py）"
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$archive = Join-Path $dist "AuditTools-windows-x64-$editionTag-$stamp.zip"
if (Test-Path -LiteralPath $archive) {
    throw "Package already exists: $archive"
}

$stage = Join-Path ([IO.Path]::GetTempPath()) ('audit-package-' + [Guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    $stageConfig = Join-Path $stage 'config-real'
    New-Item -ItemType Directory -Path $stageConfig -Force | Out-Null
    $stageWorker = Join-Path $stage 'libexec\audittools\openpyxl\worker'
    New-Item -ItemType Directory -Path $stageWorker -Force | Out-Null

    Copy-Item -LiteralPath (Join-Path $dist 'AuditTools.exe') -Destination $stage
    Copy-Item -LiteralPath (Join-Path $root 'crates\excel-editor\python\openpyxl_worker.py') `
        -Destination (Join-Path $stageWorker 'openpyxl_worker.py')
    Copy-Item -LiteralPath $guide -Destination (Join-Path $stage '使用说明.docx')
    foreach ($name in $includedConfigs) {
        Copy-Item -LiteralPath (Join-Path $configSource $name) -Destination $stageConfig
    }
    Copy-Item -LiteralPath $externalSource -Destination $stageConfig
    Copy-Item -LiteralPath $templateDirSource -Destination $stageConfig -Recurse

    if ($Edition -eq 's1s2') {
        Set-Content -LiteralPath (Join-Path $stage 'edition-s1s2') -Value 's1s2' -NoNewline -Encoding utf8
    }

    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $archive
    $sizeMb = [math]::Round((Get-Item -LiteralPath $archive).Length / 1MB, 1)
    Write-Host "Package created: $archive ($sizeMb MB)"
    Write-Host "Edition: $editionTag | configs: $($includedConfigs.Count) | guide: $guideName"
}
finally {
    if (Test-Path -LiteralPath $stage) {
        Remove-Item -LiteralPath $stage -Recurse -Force
    }
}
