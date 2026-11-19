# DiskGuard 自动化测试（Windows PowerShell 5.1）
#
# 覆盖《DiskGuard 算法设计与实现规范 v1.0》第 13 节要求的 20 类场景，
# 以及真实 ST3500418AS 双报告验收（若报告存在）。
#
# 运行：
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\DiskGuard.Tests.ps1
#
# 退出码：0 全部通过；1 存在失败。

[CmdletBinding()]
param(
    [switch]$SkipRealReports,
    [switch]$KeepTemp
)

$ErrorActionPreference = 'Stop'
$script:Here = $PSScriptRoot
$script:Root = Split-Path -Parent $script:Here
$script:Pass = 0
$script:Fail = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:CaseNo = 0
$script:Spc = 16065
$script:TempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('DiskGuardTests_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
New-Item -ItemType Directory -Path $script:TempRoot -Force | Out-Null

# 载入被测脚本（点源加载不会执行主流程）
. (Join-Path $script:Root 'DiskGuard.ps1')

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Assert-Equal($Expected, $Actual, [string]$Message) {
    if ($Expected -ne $Actual) { throw ("{0}（期望 {1}，实际 {2}）" -f $Message, $Expected, $Actual) }
}
function Assert-Match([string]$Text, [string]$Pattern, [string]$Message) {
    if ($Text -notmatch $Pattern) { throw ("{0}（未匹配 /{1}/）" -f $Message, $Pattern) }
}
function Assert-NotMatch([string]$Text, [string]$Pattern, [string]$Message) {
    if ($Text -match $Pattern) { throw ("{0}（不应匹配 /{1}/）" -f $Message, $Pattern) }
}
function Test-Case([string]$Name, [scriptblock]$Body) {
    $script:CaseNo++
    try {
        & $Body
        $script:Pass++
        Write-Host ("PASS  {0}" -f $Name) -ForegroundColor Green
    } catch {
        $script:Fail++
        $script:Failures.Add(("{0} :: {1}" -f $Name, $_.Exception.Message))
        Write-Host ("FAIL  {0} :: {1}" -f $Name, $_.Exception.Message) -ForegroundColor Red
    }
}

function New-SyntheticReport {
    param(
        [string]$Path,
        [int]$Cylinders = 1200,
        [int]$Spc = 16065,
        [int]$DefaultMs = 60,
        [int]$Spread = 0,
        [hashtable]$MsMap = @{},
        [int[]]$ErrorCylinders = @(),
        [int]$LastSectors = 0,
        [string]$Model = 'HD0:SYNTHETIC(1GB)',
        [int]$RangeEnd = -1,
        [int[]]$SkipCylinders = @(),
        [int[]]$DuplicateCylinders = @(),
        [string[]]$JunkLines = @(),
        [int]$StartCylinder = 0,
        [hashtable]$SectorCountOverrides = @{}
    )
    $last = if ($LastSectors -gt 0) { $LastSectors } else { $Spc }
    $rangeEnd = if ($RangeEnd -ge 0) { $RangeEnd } else { $StartCylinder + $Cylinders - 1 }
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('DiskGenius 磁盘检测结果：')
    $lines.Add('')
    $lines.Add(('被检测磁盘：' + $Model))
    $lines.Add(('检测范围：柱面 {0} - {1} (该磁盘柱面总数：{1})' -f $StartCylinder, $rangeEnd))
    $lines.Add('完成时间：2026-01-01 00:00:00')
    $lines.Add('')
    $lines.Add(('共检测到 {0} 个坏道区域。' -f $ErrorCylinders.Count))
    $lines.Add('')
    $lines.Add('         柱面号        磁头号          扇区号         扇区数目           耗时(ms)         修复情况   当前错误信息')
    $lines.Add('')
    foreach ($c in ($ErrorCylinders | Sort-Object)) {
        $cyl = $StartCylinder + $c
        $sec = if ($c -eq ($Cylinders - 1)) { $last } else { $Spc }
        if ($SectorCountOverrides.ContainsKey($cyl)) { $sec = [int]$SectorCountOverrides[$cyl] }
        $lines.Add(('{0,12}        未检测          未检测        {1,6}         {2,8}          未处理   Unknown Error.' -f $cyl, $sec, 3000))
        $lines.Add('')
    }
    $lines.Add(('共检测到 {0} 个正常磁道。' -f ($Cylinders - $ErrorCylinders.Count)))
    $lines.Add('')
    $lines.Add('         柱面号        磁头号          扇区号         扇区数目           耗时(ms)         修复情况   当前错误信息')
    $lines.Add('')
    for ($i = 0; $i -lt $Cylinders; $i++) {
        $cyl = $StartCylinder + $i
        if ($SkipCylinders -contains $cyl) { continue }
        if ($ErrorCylinders -contains $cyl) { continue }
        $isMapped = $MsMap.ContainsKey($cyl)
        $v = if ($isMapped) { [int]$MsMap[$cyl] } else { $DefaultMs }
        if ($Spread -gt 0 -and -not $isMapped) { $v += [int](($cyl * 7919) % ($Spread + 1)) }
        $sec = if ($i -eq ($Cylinders - 1)) { $last } else { $Spc }
        if ($SectorCountOverrides.ContainsKey($cyl)) { $sec = [int]$SectorCountOverrides[$cyl] }
        $lines.Add(('{0,12}        未检测          未检测        {1,6}         {2,8}          未处理   ' -f $cyl, $sec, $v))
        $lines.Add('')
    }
    foreach ($d in $DuplicateCylinders) {
        $lines.Add(('{0,12}        未检测          未检测        {1,6}         {2,8}          未处理   ' -f $d, $Spc, $DefaultMs))
    }
    foreach ($j in $JunkLines) { $lines.Add($j) }
    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), [System.Text.Encoding]::Unicode)
}

function Get-RunOutput([string]$Dir) {
    $cylPath = Join-Path $Dir 'Cylinder_Comparison.csv'
    $cyl = @()
    if (Test-Path -LiteralPath $cylPath) { $cyl = @(Import-Csv -LiteralPath $cylPath -Encoding UTF8) }
    $partPath = Join-Path $Dir 'Partition_Scheme.csv'
    $parts = @()
    if (Test-Path -LiteralPath $partPath) { $parts = @(Import-Csv -LiteralPath $partPath -Encoding UTF8) }
    return [pscustomobject]@{
        Dir      = $Dir
        Scan     = @(Import-Csv -LiteralPath (Join-Path $Dir 'Scan_Comparison.csv') -Encoding UTF8)
        Details  = @(Import-Csv -LiteralPath (Join-Path $Dir 'Anomaly_Details.csv') -Encoding UTF8)
        Clusters = @(Import-Csv -LiteralPath (Join-Path $Dir 'Anomaly_Clusters.csv') -Encoding UTF8)
        Ranges   = @(Import-Csv -LiteralPath (Join-Path $Dir 'Isolation_Ranges.csv') -Encoding UTF8)
        Parts    = $parts
        Cyl      = $cyl
        Plan     = (Get-Content -LiteralPath (Join-Path $Dir 'Partition_Plan.txt') -Raw -Encoding UTF8)
        Log      = (Get-Content -LiteralPath (Join-Path $Dir 'Analysis.log') -Raw -Encoding UTF8)
    }
}

function Invoke-TestRun {
    param([string]$R1, [string]$R2, [string]$Tag, [string]$ConfigPath, [switch]$NoAccelerator)
    $outRoot = Join-Path $script:TempRoot $Tag
    New-Item -ItemType Directory -Path $outRoot -Force | Out-Null
    $splat = @{ Report1 = $R1; OutputRoot = $outRoot; NoBrowser = $true; NoPause = $true }
    if ($NoAccelerator) { $splat.NoAccelerator = $true }
    if ($ConfigPath) { $splat.ConfigPath = $ConfigPath }
    if ($R2) { $splat.Report2 = $R2 } else { $splat.SingleScan = $true }
    $dir = Invoke-DiskGuard @splat 6>$null
    return (Get-RunOutput $dir)
}

Write-Host '==============================================' -ForegroundColor Cyan
Write-Host ' DiskGuard 自动化测试' -ForegroundColor Cyan
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host ("临时目录：{0}" -f $script:TempRoot)

# ---------------------------------------------------------------------------
# 01 全部正常、分布稳定
# ---------------------------------------------------------------------------
Test-Case '01 全盘正常稳定分布' {
    $r = Join-Path $script:TempRoot 'r01.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -Spread 10
    $run = Invoke-TestRun -R1 $r -Tag 't01'
    Assert-Equal 0 ([int]$run.Scan[0].ErrorRecords) '错误记录数应为 0'
    Assert-Equal 0 $run.Ranges.Count '稳定分布不应生成隔离区间'
    Assert-Equal 0 $run.Details.Count '不应有异常明细'
}

# ---------------------------------------------------------------------------
# 02 一个孤立极端慢点
# ---------------------------------------------------------------------------
Test-Case '02 孤立极端慢点' {
    $r = Join-Path $script:TempRoot 'r02.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -Spread 10 -MsMap @{ 600 = 2000 }
    $run = Invoke-TestRun -R1 $r -Tag 't02'
    Assert-Equal 1 $run.Ranges.Count '应生成 1 个隔离区间'
    Assert-Equal 'D' $run.Ranges[0].Risk '单次极端应为 D'
    $hit = @($run.Clusters | Where-Object { [int]$_.StartCylinder -eq 600 })
    Assert-Equal 1 $hit.Count '应能追溯到柱面 600 的异常簇'
}

# ---------------------------------------------------------------------------
# 03 一个孤立轻度慢点（观察，不隔离）
# ---------------------------------------------------------------------------
Test-Case '03 孤立轻度慢点不自动隔离' {
    $r = Join-Path $script:TempRoot 'r03.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 95 }
    $run = Invoke-TestRun -R1 $r -Tag 't03'
    Assert-Equal 0 $run.Ranges.Count '孤立轻度慢点不应隔离'
    Assert-True ($run.Clusters.Count -ge 1) '应记录观察簇'
    Assert-Equal 'F' $run.Clusters[0].Risk '应为 F 观察'
}

# ---------------------------------------------------------------------------
# 04 同一柱面两次扫描均异常
# ---------------------------------------------------------------------------
Test-Case '04 两次扫描同一柱面均极端' {
    $r1 = Join-Path $script:TempRoot 'r04a.txt'
    $r2 = Join-Path $script:TempRoot 'r04b.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60 -Spread 10 -MsMap @{ 600 = 2000 }
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60 -Spread 10 -MsMap @{ 600 = 2100 }
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't04'
    Assert-Equal 1 $run.Ranges.Count '应生成 1 个隔离区间'
    Assert-Equal 'B' $run.Ranges[0].Risk '两次极端应为 B'
}

# ---------------------------------------------------------------------------
# 05 单次异常、复测正常
# ---------------------------------------------------------------------------
Test-Case '05 单次极端、复测正常仍隔离' {
    $r1 = Join-Path $script:TempRoot 'r05a.txt'
    $r2 = Join-Path $script:TempRoot 'r05b.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60 -Spread 10 -MsMap @{ 600 = 2000 }
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60 -Spread 10
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't05'
    Assert-Equal 1 $run.Ranges.Count '复测正常不得抵消首次极端异常'
    Assert-Equal 'D' $run.Ranges[0].Risk '应为 D'
}

# ---------------------------------------------------------------------------
# 06 连续慢区
# ---------------------------------------------------------------------------
Test-Case '06 连续慢区（41 柱面）' {
    $r = Join-Path $script:TempRoot 'r06.txt'
    $map = @{}
    for ($i = 500; $i -le 540; $i++) { $map[$i] = 400 }
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't06'
    Assert-True ($run.Ranges.Count -ge 1) '连续慢区应生成隔离区间'
    $hit = @($run.Clusters | Where-Object { $_.Kind -eq 'Background' })
    Assert-True ($hit.Count -ge 1) '应识别为区域整体偏慢'
    Assert-Match $run.Plan '柱面范围' '计划中应有隔离区间表'
}

# ---------------------------------------------------------------------------
# 07 中间夹杂正常柱面的密集慢区
# ---------------------------------------------------------------------------
Test-Case '07 间隔式密集慢区' {
    $r = Join-Path $script:TempRoot 'r07.txt'
    $map = @{}
    for ($i = 500; $i -le 540; $i += 2) { $map[$i] = 400 }
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't07'
    Assert-True ($run.Ranges.Count -ge 1) '密集慢区应生成隔离区间'
    Assert-Equal 1 $run.Ranges.Count '间隔 2 柱面的慢点应聚成一个隔离区间'
}

# ---------------------------------------------------------------------------
# 08 两个相距很远的独立异常簇
# ---------------------------------------------------------------------------
Test-Case '08 两个独立异常簇' {
    $r = Join-Path $script:TempRoot 'r08.txt'
    $map = @{}
    for ($i = 300; $i -le 310; $i++) { $map[$i] = 400 }
    for ($i = 1500; $i -le 1510; $i++) { $map[$i] = 400 }
    New-SyntheticReport -Path $r -Cylinders 2000 -DefaultMs 60 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't08'
    Assert-Equal 2 $run.Ranges.Count '两个相距很远的簇应生成两个隔离区间'
}

# ---------------------------------------------------------------------------
# 09 任意一次出现明确读取错误
# ---------------------------------------------------------------------------
Test-Case '09 明确读取错误强制隔离' {
    $r1 = Join-Path $script:TempRoot 'r09a.txt'
    $r2 = Join-Path $script:TempRoot 'r09b.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60 -ErrorCylinders @(700)
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't09'
    Assert-Equal 1 ([int]$run.Scan[0].ErrorRecords) '应解析出 1 条错误记录'
    Assert-Equal 1 $run.Ranges.Count '错误必须进入强制隔离'
    Assert-Equal 'A' $run.Ranges[0].Risk '应为 A'
    Assert-Match $run.Plan '柱面 700' '计划中应列出错误柱面'
}

# ---------------------------------------------------------------------------
# 10 MAD 为零
# ---------------------------------------------------------------------------
Test-Case '10 MAD 为零的最小尺度保护' {
    $r = Join-Path $script:TempRoot 'r10.txt'
    New-SyntheticReport -Path $r -Cylinders 800 -DefaultMs 60
    $run = Invoke-TestRun -R1 $r -Tag 't10'
    Assert-Equal 'true' ([string]$run.Scan[0].ScaleGuard).ToLower() '应触发最小尺度保护'
    Assert-Equal 0 $run.Ranges.Count '全部相同延迟不应产生隔离'
}

# ---------------------------------------------------------------------------
# 11 大范围整体偏慢
# ---------------------------------------------------------------------------
Test-Case '11 大范围整体偏慢' {
    $r = Join-Path $script:TempRoot 'r11.txt'
    $map = @{}
    for ($i = 500; $i -le 1200; $i++) { $map[$i] = 300 }
    New-SyntheticReport -Path $r -Cylinders 2000 -DefaultMs 60 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't11'
    Assert-True ($run.Ranges.Count -ge 1) '大范围慢区应被识别'
    $bg = @($run.Clusters | Where-Object { $_.Kind -eq 'Background' })
    Assert-True ($bg.Count -ge 1) '应标记为区域整体偏慢'
    Assert-True ([int]$bg[0].Span -ge 500) '区域跨度应覆盖大部分慢区'
}

# ---------------------------------------------------------------------------
# 12 两次扫描异常位置略有变化、区域整体持续偏慢
# ---------------------------------------------------------------------------
Test-Case '12 区域级重复性' {
    $r1 = Join-Path $script:TempRoot 'r12a.txt'
    $r2 = Join-Path $script:TempRoot 'r12b.txt'
    $m1 = @{}; for ($i = 500; $i -le 900; $i++) { $m1[$i] = 300 }
    $m2 = @{}; for ($i = 550; $i -le 950; $i++) { $m2[$i] = 300 }
    New-SyntheticReport -Path $r1 -Cylinders 2000 -DefaultMs 60 -MsMap $m1
    New-SyntheticReport -Path $r2 -Cylinders 2000 -DefaultMs 60 -MsMap $m2
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't12'
    Assert-True ($run.Ranges.Count -ge 1) '区域级重复慢区应被隔离'
    Assert-Equal 'B' $run.Ranges[0].Risk '两次扫描均持续偏慢应为 B'
}

# ---------------------------------------------------------------------------
# 13 两次扫描记录范围不一致
# ---------------------------------------------------------------------------
Test-Case '13 双报告范围不一致禁止生成 LBA' {
    $r1 = Join-Path $script:TempRoot 'r13a.txt'
    $r2 = Join-Path $script:TempRoot 'r13b.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 2000 }
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 2000 } -RangeEnd 1198
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't13'
    Assert-Match $run.Plan '扫描范围不一致' '应报告范围不一致'
    Assert-Match $run.Plan 'LBA 映射：未验证' '范围不一致时 LBA 必须未验证'
    foreach ($rg in $run.Ranges) { Assert-True ([string]::IsNullOrEmpty($rg.LbaStart)) '不得输出 LBA 边界' }
}

# ---------------------------------------------------------------------------
# 14 柱面编号重复或缺失
# ---------------------------------------------------------------------------
Test-Case '14 柱面编号重复与缺失' {
    $r1 = Join-Path $script:TempRoot 'r14a.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60 -DuplicateCylinders @(600)
    $run = Invoke-TestRun -R1 $r1 -Tag 't14dup'
    Assert-Match $run.Plan '重复柱面' '应报告重复柱面'
    Assert-Match $run.Plan 'LBA 映射：未验证' '重复柱面时 LBA 必须未验证'

    $r2 = Join-Path $script:TempRoot 'r14b.txt'
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60 -SkipCylinders @(600)
    $run2 = Invoke-TestRun -R1 $r2 -Tag 't14miss'
    Assert-Match $run2.Plan '缺失 1 个柱面' '应报告缺失柱面'
    Assert-Match $run2.Plan 'LBA 映射：未验证' '缺失柱面时 LBA 必须未验证'
}

# ---------------------------------------------------------------------------
# 15 最后一个柱面不完整
# ---------------------------------------------------------------------------
Test-Case '15 最后一个柱面不完整' {
    $r = Join-Path $script:TempRoot 'r15.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -LastSectors 5000
    $run = Invoke-TestRun -R1 $r -Tag 't15'
    Assert-Match $run.Plan 'LBA 映射：已验证' '最后柱面不完整但容量自洽时应验证通过'
    $capText = (1199 * $script:Spc + 5000).ToString()
    Assert-Match $run.Plan ('总扇区：' + $capText) '总扇区数应包含最后柱面实际扇区'
    $cap = 1199 * $script:Spc + 5000
    $last = @($run.Ranges | Where-Object { $_.LbaEnd -ne '' } | Sort-Object { [int64]$_.LbaEnd } | Select-Object -Last 1)
    if ($last.Count -gt 0) {
        Assert-True ([int64]$last[0].LbaEnd -le ($cap - 34)) 'LBA 不得超过 GPT 尾部保留区'
    }
}

# ---------------------------------------------------------------------------
# 16 存在无法解析的行
# ---------------------------------------------------------------------------
Test-Case '16 无法解析的记录必须报告' {
    $r = Join-Path $script:TempRoot 'r16.txt'
    New-SyntheticReport -Path $r -Cylinders 800 -DefaultMs 60 -JunkLines @('这是一行无法解析的扫描记录 99999')
    $run = Invoke-TestRun -R1 $r -Tag 't16'
    Assert-Match $run.Plan '无法解析' '应在结构校验中报告无法解析的行'
    Assert-Match $run.Log '未解析行' '日志应列出原始行号'
}

# ---------------------------------------------------------------------------
# 17 异常区间扩展后重叠
# ---------------------------------------------------------------------------
Test-Case '17 扩展后重叠的隔离区间应合并' {
    $r = Join-Path $script:TempRoot 'r17.txt'
    $map = @{}
    for ($i = 500; $i -le 505; $i++) { $map[$i] = 2000 }
    for ($i = 525; $i -le 530; $i++) { $map[$i] = 2000 }
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't17'
    Assert-Equal 1 $run.Ranges.Count '扩展后重叠的两个簇应合并为一个隔离区间'
    Assert-Match $run.Ranges[0].Reason '合并' '应记录合并来源'
}

# ---------------------------------------------------------------------------
# 18 隔离区间接近磁盘起止边界
# ---------------------------------------------------------------------------
Test-Case '18 起止边界裁剪与 GPT 保护' {
    $r = Join-Path $script:TempRoot 'r18.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -ErrorCylinders @(0, 1199)
    $run = Invoke-TestRun -R1 $r -Tag 't18'
    Assert-True ($run.Ranges.Count -ge 1) '边界错误应生成隔离区间'
    $cap = 1199 * $script:Spc + $script:Spc
    foreach ($rg in $run.Ranges) {
        Assert-True ([int]$rg.StartCylinder -ge 0) '柱面下界不得为负'
        Assert-True ([int]$rg.EndCylinder -le 1199) '柱面上界不得超过磁盘'
        if (-not [string]::IsNullOrEmpty($rg.LbaStart)) {
            Assert-True ([int64]$rg.LbaStart -ge 2048) 'LBA 起不得侵入 GPT 头部保留区'
            Assert-True ([int64]$rg.LbaEnd -le ($cap - 34)) 'LBA 止不得侵入 GPT 尾部保留区'
        }
    }
}

# ---------------------------------------------------------------------------
# 19 LBA 映射未经确认
# ---------------------------------------------------------------------------
Test-Case '19 手工确认模式下未确认则不给 LBA' {
    $r = Join-Path $script:TempRoot 'r19.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 2000 }
    $cfgPath = Join-Path $script:TempRoot 'config_manual.json'
    $cfg = Get-Content -LiteralPath (Join-Path $script:Root 'Config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $cfg.RequireManualLbaConfirmation = $true
    $cfg.LbaMappingConfirmed = $false
    [System.IO.File]::WriteAllText($cfgPath, ($cfg | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)
    $run = Invoke-TestRun -R1 $r -Tag 't19' -ConfigPath $cfgPath
    Assert-Match $run.Plan 'LBA 映射：未验证' '未手工确认时不得输出 LBA'
    foreach ($rg in $run.Ranges) { Assert-True ([string]::IsNullOrEmpty($rg.LbaStart)) '不得输出 LBA 边界' }
}

# ---------------------------------------------------------------------------
# 20 GPT 首尾空间与 1 MiB 对齐
# ---------------------------------------------------------------------------
Test-Case '20 1 MiB 对齐与 GPT 保留区' {
    $r = Join-Path $script:TempRoot 'r20.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 2000 }
    $run = Invoke-TestRun -R1 $r -Tag 't20'
    Assert-Match $run.Plan 'LBA 映射：已验证' '正常报告应验证通过'
    $cap = 1200 * $script:Spc
    $any = $false
    foreach ($rg in $run.Ranges) {
        if ([string]::IsNullOrEmpty($rg.LbaStart)) { continue }
        $any = $true
        Assert-Equal 0 ([int64]$rg.LbaStart % 2048) 'LBA 起必须 1 MiB 对齐'
        Assert-Equal 2047 ([int64]$rg.LbaEnd % 2048) 'LBA 止必须落在 1 MiB 边界前'
        Assert-True ([int64]$rg.LbaStart -ge 2048) 'LBA 起不得小于 2048'
        Assert-True ([int64]$rg.LbaEnd -le ($cap - 34)) 'LBA 止不得超过容量减 34'
    }
    Assert-True $any '应至少输出一个 LBA 区间'
}

# ---------------------------------------------------------------------------
# 21 分区方案参数（起始柱面/磁头/扇区/起始扇区号/总扇区数）
# ---------------------------------------------------------------------------
Test-Case '21 分区方案参数（CHS / LBA / 总扇区）' {
    $r = Join-Path $script:TempRoot 'r21p.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap @{ 600 = 2000 }
    $run = Invoke-TestRun -R1 $r -Tag 't21p'
    Assert-True ($run.Parts.Count -ge 1) '应输出分区方案'
    $spc = 16065
    foreach ($p in $run.Parts) {
        Assert-True ($null -ne $p.StartCylinder -and [string]$p.StartCylinder -ne '') '缺少起始柱面'
        Assert-True ($null -ne $p.StartHead -and [string]$p.StartHead -ne '') '缺少磁头'
        Assert-True ($null -ne $p.StartSector -and [string]$p.StartSector -ne '') '缺少扇区'
        $lba = [int64]$p.StartLba
        $expect = [int64]$p.StartCylinder * $spc + [int64]$p.StartHead * 63 + ([int64]$p.StartSector - 1)
        Assert-Equal $lba $expect 'CHS 与 LBA 不一致'
        Assert-Equal ([int64]$p.TotalSectors) ([int64]$p.EndLba - $lba + 1) '总扇区数与起止 LBA 不一致'
        Assert-Equal 0 ($lba % 2048) '起始扇区号未按 1 MiB 对齐'
    }
    Assert-Match $run.Plan '起始柱面' '计划中应有分区表'
    Assert-Match $run.Plan '总扇区数' '计划中应有总扇区数'
}

# ---------------------------------------------------------------------------
# 22 加速内核与纯 PowerShell 回退结果一致
# ---------------------------------------------------------------------------
Test-Case '22 加速内核与回退实现一致' {
    Initialize-Accelerator
    $savedAccel = $script:Accelerated
    if (-not $savedAccel) {
        Write-Host '      SKIP：当前环境无法编译 C# 加速内核，跳过一致性比较' -ForegroundColor Yellow
        return
    }
    $r = Join-Path $script:TempRoot 'r21.txt'
    New-SyntheticReport -Path $r -Cylinders 400 -DefaultMs 60 -Spread 12 -MsMap @{ 100 = 900; 200 = 200; 201 = 210; 300 = 500 }
    $cfg = Get-DiskGuardConfig (Join-Path $script:Root 'Config.json')
    try {
        $script:Accelerated = $true
        $repA = Parse-Report -Path $r -Label 'A' -Config $cfg
        $repA = Add-DetectionFlags -Report $repA -Config $cfg
        $script:Accelerated = $false
        $repB = Parse-Report -Path $r -Label 'B' -Config $cfg
        $repB = Add-DetectionFlags -Report $repB -Config $cfg
    } finally {
        $script:Accelerated = $savedAccel
    }
    for ($i = 0; $i -lt $repA.Count; $i++) {
        if ($repA.Flags.Local[$i] -ne $repB.Flags.Local[$i]) { throw ("柱面 {0} 的局部异常标记不一致" -f $repA.Cylinder[$i]) }
        if ([Math]::Abs($repA.Flags.MedianSmall[$i] - $repB.Flags.MedianSmall[$i]) -gt 0.001) { throw ("柱面 {0} 的局部中位数不一致" -f $repA.Cylinder[$i]) }
    }
}

# ---------------------------------------------------------------------------
# 23 风险策略 strict / errors-only
# ---------------------------------------------------------------------------
Test-Case '23 风险策略 strict 与 errors-only' {
    $r = Join-Path $script:TempRoot 'r23p.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -MsMap @{ 300 = 2000 } -ErrorCylinders @(800)
    $cfg = Get-Content -LiteralPath (Join-Path $script:Root 'Config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $cfg.RiskProfile = 'strict'
    $cfgPath = Join-Path $script:TempRoot 'config_strict.json'
    [System.IO.File]::WriteAllText($cfgPath, ($cfg | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)
    $strict = Invoke-TestRun -R1 $r -Tag 't23s' -ConfigPath $cfgPath
    Assert-Equal 2 $strict.Ranges.Count 'strict 应同时隔离错误与极端点'
    $cfg.RiskProfile = 'errors-only'
    $cfgPath2 = Join-Path $script:TempRoot 'config_errors.json'
    [System.IO.File]::WriteAllText($cfgPath2, ($cfg | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)
    $eo = Invoke-TestRun -R1 $r -Tag 't23e' -ConfigPath $cfgPath2
    Assert-Equal 1 $eo.Ranges.Count 'errors-only 只隔离明确错误'
    Assert-Match $eo.Plan '风险策略：errors-only' '应说明当前风险策略'
}


# ---------------------------------------------------------------------------
# 25 LBA：最后柱面完整
# ---------------------------------------------------------------------------
Test-Case '25 LBA：最后柱面完整也应验证通过' {
    $r = Join-Path $script:TempRoot 'r25.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -Spread 8
    $run = Invoke-TestRun -R1 $r -Tag 't25'
    Assert-Match $run.Plan 'LBA 映射：已验证' '完整最后柱面不应阻止映射'
    Assert-Match $run.Plan '完整柱面' '证据应说明最后柱面为完整柱面'
}

# ---------------------------------------------------------------------------
# 26 LBA：非末柱面扇区数异常
# ---------------------------------------------------------------------------
Test-Case '26 LBA：非末柱面扇区数异常必须拒绝' {
    $r = Join-Path $script:TempRoot 'r26.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -SectorCountOverrides @{ 600 = 16000 }
    $run = Invoke-TestRun -R1 $r -Tag 't26'
    Assert-Match $run.Plan 'LBA 映射：未验证' '非末柱面扇区数异常时不得输出 LBA'
    Assert-Match $run.Plan '非末柱面扇区数一致' '应列出逐条扇区数证据'
}

# ---------------------------------------------------------------------------
# 27 LBA：声明末端与实际末柱面不一致
# ---------------------------------------------------------------------------
Test-Case '27 LBA：声明末端与实际末柱面不一致必须拒绝' {
    $r = Join-Path $script:TempRoot 'r27.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -RangeEnd 1300
    $run = Invoke-TestRun -R1 $r -Tag 't27'
    Assert-Match $run.Plan 'LBA 映射：未验证' 'RangeEnd 不一致时不得输出 LBA'
    Assert-Match $run.Plan '声明扫描末端' '校验应指出末端不一致'
}

# ---------------------------------------------------------------------------
# 28 LBA：报告尾部缺失
# ---------------------------------------------------------------------------
Test-Case '28 LBA：报告尾部缺失必须拒绝' {
    $r = Join-Path $script:TempRoot 'r28.txt'
    $skip = @(1100..1199)
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -SkipCylinders $skip
    $run = Invoke-TestRun -R1 $r -Tag 't28'
    Assert-Match $run.Plan 'LBA 映射：未验证' '尾部缺失时不得输出 LBA'
    Assert-Match $run.Plan '实际解析 1100 条记录' '校验应指出记录数不足'
}

# ---------------------------------------------------------------------------
# 29 区域慢区：稀疏异常点不得合并成连续慢区
# ---------------------------------------------------------------------------
Test-Case '29 区域慢区：稀疏异常点不得合并' {
    $r = Join-Path $script:TempRoot 'r29.txt'
    $map = @{ 100 = 120; 110 = 120; 120 = 120 }
    for ($c = 600; $c -le 980; $c += 20) { $map[$c] = 150 }
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -Spread 40 -MsMap $map
    $run = Invoke-TestRun -R1 $r -Tag 't29'
    $bg = @($run.Clusters | Where-Object { $_.Kind -eq 'Background' })
    Assert-Equal 0 $bg.Count '稀疏点不应形成背景慢区'
    Assert-Equal 0 $run.Ranges.Count '稀疏点不应产生隔离区间'
}

# ---------------------------------------------------------------------------
# 30 区域慢区：仅第二次扫描出现的慢区
# ---------------------------------------------------------------------------
Test-Case '30 区域慢区：仅第二次扫描出现也要识别' {
    $r1 = Join-Path $script:TempRoot 'r30a.txt'
    $r2 = Join-Path $script:TempRoot 'r30b.txt'
    $map2 = @{}
    for ($c = 500; $c -le 900; $c++) { $map2[$c] = 300 }
    New-SyntheticReport -Path $r1 -Cylinders 2000 -DefaultMs 60 -Spread 10
    New-SyntheticReport -Path $r2 -Cylinders 2000 -DefaultMs 60 -Spread 10 -MsMap $map2
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't30'
    Assert-True ($run.Ranges.Count -ge 1) 'R2 独有慢区应产生隔离区间'
    $bg = @($run.Clusters | Where-Object { $_.Kind -eq 'Background' })
    Assert-True ($bg.Count -ge 1) '应识别为区域慢区'
}

# ---------------------------------------------------------------------------
# 31 非零起始柱面
# ---------------------------------------------------------------------------
Test-Case '31 非零起始柱面：拒绝 LBA 且不崩溃' {
    $r = Join-Path $script:TempRoot 'r31.txt'
    New-SyntheticReport -Path $r -Cylinders 600 -StartCylinder 1000 -DefaultMs 60 -MsMap @{ 1200 = 2000 }
    $run = Invoke-TestRun -R1 $r -Tag 't31'
    Assert-Match $run.Plan 'LBA 映射：未验证' '非零起始柱面不得输出 LBA'
    Assert-Match $run.Plan '未从柱面 0 开始' '校验应指出起始柱面非 0'
}

# ---------------------------------------------------------------------------
# 32 双报告记录数不一致
# ---------------------------------------------------------------------------
Test-Case '32 双报告记录数不一致：禁止正式 LBA' {
    $r1 = Join-Path $script:TempRoot 'r32a.txt'
    $r2 = Join-Path $script:TempRoot 'r32b.txt'
    New-SyntheticReport -Path $r1 -Cylinders 1200 -DefaultMs 60
    New-SyntheticReport -Path $r2 -Cylinders 1200 -DefaultMs 60 -SkipCylinders @(600)
    $run = Invoke-TestRun -R1 $r1 -R2 $r2 -Tag 't32'
    Assert-Match $run.Plan 'LBA 映射：未验证' '记录数不一致时不得输出 LBA'
    Assert-Match $run.Plan '记录数不同' '应说明记录数不一致'
    Assert-Match $run.Plan '已跳过局部异常检测' '柱面不连续时应跳过按索引的局部检测'
}

# ---------------------------------------------------------------------------
# 33 纯 PowerShell 回退
# ---------------------------------------------------------------------------
Test-Case '33 纯 PowerShell 回退可运行' {
    $r = Join-Path $script:TempRoot 'r33.txt'
    New-SyntheticReport -Path $r -Cylinders 400 -DefaultMs 60 -Spread 12 -MsMap @{ 200 = 900 }
    $run = Invoke-TestRun -R1 $r -Tag 't33' -NoAccelerator
    Assert-Match $run.Plan '分区方案' '回退路径应生成分区方案'
    Assert-True ($run.Ranges.Count -ge 1) '回退路径应识别异常'
}

# ---------------------------------------------------------------------------
# 34 SPC 未知时不除零
# ---------------------------------------------------------------------------
Test-Case '34 SPC 未知时隔离区间不除零' {
    $cfg = Get-DiskGuardConfig (Join-Path $script:Root 'Config.json')
    $mapping = [pscustomobject]@{ Verified = $false; SectorsPerCylinder = 0; CapacitySectors = 0 }
    $cluster = [pscustomobject]@{
        Isolated = $true; GuardMiB = 512.0; StartCylinder = 100; EndCylinder = 120
        Risk = 'C'; RiskReason = 'test'; MaxMs1 = 500.0; MaxMs2 = 500.0
        RegionP951 = 400.0; RegionP952 = 400.0; PointCount = 21; ErrorPoints = 0
        ErrorCylinders = @(); Kind = 'Anomaly'; Members = @(100..120)
    }
    $iso = Build-IsolationRanges -Clusters @($cluster) -Mapping $mapping -Config $cfg -LastCylinder 1199
    Assert-Equal 1 $iso.Ranges.Count '应仍输出柱面范围'
    Assert-Equal 100 ([int64]$iso.Ranges[0].StartCylinder) 'SPC 未知时保护带应为 0'
}


# ---------------------------------------------------------------------------
# 35 Linux：验证通过时生成分区与 LVM 脚本
# ---------------------------------------------------------------------------
Test-Case '35 Linux：生成分区与 LVM 脚本' {
    $r = Join-Path $script:TempRoot 'r35.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -Spread 8 -MsMap @{ 600 = 2000 }
    $run = Invoke-TestRun -R1 $r -Tag 't35'
    $sh = Join-Path $run.Dir 'Linux_Setup.sh'
    Assert-True (Test-Path -LiteralPath $sh) '应生成 Linux_Setup.sh'
    $text = Get-Content -LiteralPath $sh -Raw -Encoding UTF8
    Assert-Match $text 'sgdisk -n 1:' '应包含 sgdisk 分区命令'
    Assert-Match $text 'pvcreate' '应包含 pvcreate'
    Assert-Match $text 'vgcreate "\$VG"' '应包含 vgcreate'
    Assert-Match $text 'lvcreate -l 100%FREE' '应包含 lvcreate'
    Assert-Match $text 'mkfs\.ext4' '应包含 mkfs.ext4'
    Assert-Match $text 'SERIAL=' '应包含序列号读取'
    Assert-Match $text 'vg_' '应包含 vg_ 前缀'
    $bytes = [System.IO.File]::ReadAllBytes($sh)
    Assert-True (($bytes[0] -ne 0xEF) -or ($bytes[1] -ne 0xBB) -or ($bytes[2] -ne 0xBF)) '脚本不应带 BOM'
    Assert-Match $run.Plan 'Linux_Setup.sh' '计划中应提到脚本'
}

# ---------------------------------------------------------------------------
# 36 Linux：LBA 未验证时不生成脚本
# ---------------------------------------------------------------------------
Test-Case '36 Linux：LBA 未验证时不生成脚本' {
    $r = Join-Path $script:TempRoot 'r36.txt'
    New-SyntheticReport -Path $r -Cylinders 1200 -DefaultMs 60 -RangeEnd 1300
    $run = Invoke-TestRun -R1 $r -Tag 't36'
    $sh = Join-Path $run.Dir 'Linux_Setup.sh'
    Assert-True (-not (Test-Path -LiteralPath $sh)) '未验证时不应生成脚本'
    Assert-Match $run.Plan '未生成 Linux_Setup.sh' '计划中应说明未生成'
}

# ---------------------------------------------------------------------------
# 24 真实 ST3500418AS 双报告验收
# ---------------------------------------------------------------------------
if (-not $SkipRealReports) {
    $real1 = Join-Path $script:Root 'BadTrack Report ST3500418AS(466GB).txt'
    $real2 = Join-Path $script:Root 'BadTrack Report ST3500418AS(466GB)2.txt'
    if ((Test-Path -LiteralPath $real1) -and (Test-Path -LiteralPath $real2)) {
        Test-Case '24 真实报告：性能与断言' {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $run = Invoke-TestRun -R1 $real1 -R2 $real2 -Tag 't22'
            $sw.Stop()
            Assert-True ($sw.Elapsed.TotalSeconds -lt 60) ("总耗时 {0:N1}s 超过 60s" -f $sw.Elapsed.TotalSeconds)
            Assert-Equal 3 ([int]$run.Scan[0].ErrorRecords) '扫描1 应有 3 条错误'
            Assert-Equal 3 ([int]$run.Scan[1].ErrorRecords) '扫描2 应有 3 条错误'
            Assert-Equal 60802 ([int]$run.Scan[0].Records) '记录数应为 60802'
            Assert-Match $run.Plan '总扇区：976773168' '总扇区数应为 976,773,168'
            Assert-Match $run.Plan 'LBA 映射：已验证' '真实报告映射应验证通过'
            Assert-True ($run.Ranges.Count -le 40) '隔离区间数量应可控（≤40）'
            Assert-True ($run.Parts.Count -ge 1) '应输出分区方案'
            $firstPart = @($run.Parts | Where-Object { $_.Recommended -eq 'true' } | Select-Object -First 1)
            Assert-True ($firstPart.Count -eq 1) '应至少有一个推荐分区'
            Assert-True ([int64]$firstPart[0].StartLba -ge 2048) '首个分区不得侵入 GPT 头部保留区'
            foreach ($c in 16, 4708, 13585) {
                $d = @($run.Details | Where-Object { [int]$_.Cylinder -eq $c })
                Assert-Equal 1 $d.Count ("柱面 {0} 应出现在异常明细中" -f $c)
                Assert-Equal 'true' $d[0].Error1 ("柱面 {0} 应标记为错误" -f $c)
            }
            foreach ($c in 40, 4186, 4709, 12019, 39952, 39953, 40737, 13586) {
                $d = @($run.Details | Where-Object { [int]$_.Cylinder -eq $c })
                Assert-Equal 1 $d.Count ("参考慢点柱面 {0} 应出现在异常明细中" -f $c)
            }
            Write-Host ("      真实报告耗时 {0:N1}s；隔离区间 {1} 个" -f $sw.Elapsed.TotalSeconds, $run.Ranges.Count) -ForegroundColor DarkGray
        }
    } else {
        Write-Host 'SKIP  24 真实报告验收（未找到 ST3500418AS 报告）' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# 汇总
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host '==============================================' -ForegroundColor Cyan
Write-Host (" 通过 {0} / 失败 {1}" -f $script:Pass, $script:Fail) -ForegroundColor $(if ($script:Fail -eq 0) { 'Green' } else { 'Red' })
if ($script:Fail -gt 0) {
    foreach ($f in $script:Failures) { Write-Host ("  - " + $f) -ForegroundColor Red }
}
Write-Host (" 临时目录：{0}" -f $script:TempRoot)
Write-Host '==============================================' -ForegroundColor Cyan

if (-not $KeepTemp) {
    try { Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
if ($script:Fail -gt 0) { exit 1 } else { exit 0 }