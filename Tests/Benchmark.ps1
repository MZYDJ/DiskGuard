# DiskGuard 性能基准（Windows PowerShell 5.1）
#
# 用法：
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Benchmark.ps1
#   powershell.exe ... -Runs 5 -IncludeFallback
#
# 默认读取仓库根目录下的 ST3500418AS 双报告；报告不存在时直接报错。

[CmdletBinding()]
param(
    [string]$Report1 = '',
    [string]$Report2 = '',
    [int]$Runs = 3,
    [switch]$IncludeFallback
)

$ErrorActionPreference = 'Stop'
$script:Root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Report1)) {
    $Report1 = Join-Path $script:Root 'BadTrack Report ST3500418AS(466GB).txt'
}
if ([string]::IsNullOrWhiteSpace($Report2)) {
    $Report2 = Join-Path $script:Root 'BadTrack Report ST3500418AS(466GB)2.txt'
}
if (-not (Test-Path -LiteralPath $Report1)) { throw ("找不到报告：{0}" -f $Report1) }
if (-not (Test-Path -LiteralPath $Report2)) { $Report2 = $null }

# 点源 DiskGuard.ps1 会重新绑定其 param 块（Report1/Report2 被清空），
# 所以先把路径存到独立变量。
$script:BenchR1 = $Report1
$script:BenchR2 = $Report2

. (Join-Path $script:Root 'DiskGuard.ps1')

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('DiskGuardBench_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $temp -Force | Out-Null

function Invoke-BenchRun([string]$Tag, [switch]$NoAccelerator) {
    $outRoot = Join-Path $temp $Tag
    New-Item -ItemType Directory -Path $outRoot -Force | Out-Null
    $splat = @{
        Report1    = $script:BenchR1
        OutputRoot = $outRoot
        NoBrowser  = $true
        NoPause    = $true
    }
    if ($script:BenchR2) { $splat.Report2 = $script:BenchR2 } else { $splat.SingleScan = $true }
    if ($NoAccelerator) { $splat.NoAccelerator = $true }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $dir = Invoke-DiskGuard @splat 6>$null
    $sw.Stop()
    $ranges = @(Import-Csv -LiteralPath (Join-Path $dir 'Isolation_Ranges.csv') -Encoding UTF8)
    $scan = @(Import-Csv -LiteralPath (Join-Path $dir 'Scan_Comparison.csv') -Encoding UTF8)
    return [pscustomobject]@{
        Tag        = $Tag
        Seconds    = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
        Records1   = if ($scan.Count -ge 1) { [int]$scan[0].Records } else { 0 }
        Records2   = if ($scan.Count -ge 2) { [int]$scan[1].Records } else { 0 }
        Ranges     = $ranges.Count
        Accelerated = [bool]$script:Accelerated
        Dir        = $dir
    }
}

$results = New-Object 'System.Collections.Generic.List[object]'
for ($i = 1; $i -le $Runs; $i++) {
    Write-Host ("加速内核 第 {0}/{1} 次……" -f $i, $Runs) -ForegroundColor Cyan
    $results.Add((Invoke-BenchRun -Tag ('acc{0}' -f $i)))
}
if ($IncludeFallback) {
    Write-Host '纯 PowerShell 回退（较慢）……' -ForegroundColor Yellow
    $results.Add((Invoke-BenchRun -Tag 'slow' -NoAccelerator))
}

Write-Host ''
Write-Host '模式        耗时(s)  记录1   记录2   隔离区  加速' -ForegroundColor Cyan
foreach ($r in $results) {
    Write-Host ("{0,-10} {1,7:N2}  {2,6}  {3,6}  {4,6}  {5}" -f $r.Tag, $r.Seconds, $r.Records1, $r.Records2, $r.Ranges, $(if ($r.Accelerated) { '是' } else { '否' }))
}
Write-Host ''
Write-Host ("输出目录：{0}" -f $temp)