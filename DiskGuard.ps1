# DiskGuard v2.0.0
# DiskGenius 全盘坏道扫描报告分析器（Windows PowerShell 5.1，严格只读）
#
# 设计依据：DiskGuard_算法设计与实现规范_v1.0.md
# 处理链：解析与结构校验 -> 单次扫描统计基线 -> 全盘异常 -> 局部异常
#         -> 双报告对照 -> 区域分析 -> 一维聚类 -> 风险分级
#         -> 安全隔离区间 -> LBA 映射验证 -> 分区规划
#
# 本程序只读取 TXT 报告并写出分析文件，不会打开物理磁盘，不会修改分区表。

[CmdletBinding()]
param(
    [string]$Report1,
    [string]$Report2,
    [string]$OutputRoot,
    [switch]$NoBrowser,
    [switch]$NoPause,
    [switch]$SingleScan,
    [switch]$NoAccelerator
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

$script:Version = '2.0.0'
$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:RunClock = New-Object System.Diagnostics.Stopwatch
$script:StageClock = New-Object System.Diagnostics.Stopwatch
$script:FailureLog = $null
$script:Accelerated = $false
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:MiB = 1048576.0

# ---------------------------------------------------------------------------
# 基础输出
# ---------------------------------------------------------------------------
function Add-Log([string]$Text) {
    $script:LogLines.Add(('{0:HH:mm:ss.fff}  {1}' -f (Get-Date), $Text))
}
function Write-Info([string]$Message) {
    Write-Host "[DiskGuard] $Message" -ForegroundColor Cyan
    Add-Log $Message
}
function Write-Warn([string]$Message) {
    Write-Host "[警告] $Message" -ForegroundColor Yellow
    Add-Log ('WARN  ' + $Message)
}
function Write-Fail([string]$Message) {
    Write-Host "[错误] $Message" -ForegroundColor Red
    Add-Log ('ERROR ' + $Message)
}
function Format-Num($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        $d = [double]$Value
        if ([double]::IsNaN($d)) { return '' }
        return $d.ToString('0.###', $script:Inv)
    }
    return [string]$Value
}
function ConvertTo-CsvText($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal]) {
        return ([double]$Value).ToString('0.###', $script:Inv)
    }
    $s = [string]$Value
    if ($s.IndexOfAny([char[]]@(',', '"', "`r", "`n")) -ge 0) {
        return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}
function Get-DiskGuardConfig([string]$Path) {
    $cfg = [pscustomobject]@{
        LogicalSectorBytes            = 512
        MinRobustSigmaMs              = 5.0
        GlobalK                       = 4.0
        ExtremeK                      = 10.0
        ExtremeRatioMin               = 10.0
        StrongRatioMin                = 6.0
        LocalWindowSmall              = 50
        LocalWindowLarge              = 200
        LocalK                        = 4.0
        LocalMinSigmaMs               = 1.0
        LocalRatioMin                 = 1.5
        LocalMinDeltaMs               = 100.0
        LocalMinSamples               = 8
        TightClusterCylinders         = 10
        MaxClusterSpan                = 200
        MinDensePoints                = 3
        ContextCylinders              = 10
        RegionSlowRatio               = 1.5
        MinBackgroundSpan             = 20
        GuardBackgroundMiB            = 512
        RiskProfile                   = 'strict'
        MinPartitionGiB               = 1.0
        MinPartitionPercent           = 20.0
        HeadsPerCylinder              = 255
        SectorsPerTrack               = 63
        WriteCylinderComparison       = $false
        GuardErrorMiB                 = 1024
        GuardExtremeMiB               = 512
        GuardRepeatedMiB              = 256
        GuardSingleMiB                = 256
        DiskCapacitySectorsOverride   = $null
        SectorsPerCylinderOverride    = $null
        LbaMappingConfirmed           = $false
        RequireManualLbaConfirmation  = $false
    }
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $user = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $user.PSObject.Properties) {
                if ($null -ne $cfg.PSObject.Properties[$p.Name]) {
                    $cfg.($p.Name) = $p.Value
                } else {
                    Write-Warn "Config.json 中存在未知参数：$($p.Name)（已忽略）"
                }
            }
        } catch {
            Write-Warn ("Config.json 读取失败，使用内置默认值：{0}" -f $_.Exception.Message)
        }
    }
    return $cfg
}

# ---------------------------------------------------------------------------
# 报告选择与编码
# ---------------------------------------------------------------------------
function Select-Report([string]$Title, [bool]$Required) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = $Title
        $dlg.Filter = 'DiskGenius 报告 (*.txt)|*.txt|所有文件 (*.*)|*.*'
        $dlg.Multiselect = $false
        $dlg.CheckFileExists = $true
        $dlg.RestoreDirectory = $true
        $result = $dlg.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
        if ($Required) { throw '必须选择第一次扫描报告。' }
        return $null
    } catch {
        if ($_.Exception.Message -like '*必须选择*') { throw }
        Write-Warn '当前环境无法使用 Windows 文件选择窗口，改用命令行输入路径。'
        if ($Required) {
            $p = Read-Host '请输入第一次扫描 TXT 报告的完整路径'
            if ([string]::IsNullOrWhiteSpace($p)) { throw '未提供第一次扫描报告。' }
            return $p.Trim('"')
        }
        $p = Read-Host '请输入第二次扫描 TXT 报告路径（直接回车跳过）'
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        return $p.Trim('"')
    }
}
function Get-TextEncoding([string]$Path) {
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $buf = New-Object byte[] 65536
        $n = $fs.Read($buf, 0, 65536)
        if ($n -ge 2 -and $buf[0] -eq 0xFF -and $buf[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
        if ($n -ge 2 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
        if ($n -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8 }
        $oddZeros = 0; $evenZeros = 0
        for ($i = 0; $i -lt $n; $i++) {
            if ($buf[$i] -eq 0) { if (($i % 2) -eq 0) { $evenZeros++ } else { $oddZeros++ } }
        }
        if ($oddZeros -gt [int]($n / 4) -and $evenZeros -lt 2) { return [System.Text.Encoding]::Unicode }
        try {
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            $probe = [Math]::Max(0, $n - 4)
            if ($probe -gt 0) { [void]$strict.GetString($buf, 0, $probe) }
            return [System.Text.Encoding]::UTF8
        } catch {
            return [System.Text.Encoding]::Default
        }
    } finally {
        $fs.Dispose()
    }
}

# ---------------------------------------------------------------------------
# 统计
# ---------------------------------------------------------------------------
function Get-PercentileSorted([double[]]$Sorted, [double]$P) {
    $n = $Sorted.Length
    if ($n -eq 0) { return 0.0 }
    if ($n -eq 1) { return [double]$Sorted[0] }
    $idx = ($n - 1) * $P
    $lo = [int][Math]::Floor($idx)
    $hi = [int][Math]::Ceiling($idx)
    if ($lo -eq $hi) { return [double]$Sorted[$lo] }
    return [double]$Sorted[$lo] + (([double]$Sorted[$hi] - [double]$Sorted[$lo]) * ($idx - $lo))
}
function Get-Statistics {
    param([double[]]$ReadMs, [bool[]]$IsError, $Config)

    $n = $ReadMs.Length
    $validCount = 0
    $errCount = 0
    $sum = 0.0
    $valid = New-Object double[] $n
    for ($i = 0; $i -lt $n; $i++) {
        if ($IsError[$i]) { $errCount++; continue }
        $v = $ReadMs[$i]
        if ([double]::IsNaN($v)) { continue }
        $valid[$validCount] = $v
        $validCount++
        $sum += $v
    }
    if ($validCount -eq 0) {
        return [pscustomobject]@{
            ValidCount = 0; ErrorCount = $errCount; MissingCount = ($n - $errCount)
            Min = 0.0; Max = 0.0; Mean = 0.0; Median = 0.0; MAD = 0.0; RobustSigma = 1.0
            P25 = 0.0; P50 = 0.0; P75 = 0.0; P90 = 0.0; P95 = 0.0; P97 = 0.0; P98 = 0.0
            P99 = 0.0; P995 = 0.0; P999 = 0.0
            TGlobal = 0.0; TExtremeStat = 0.0; TExtreme = 0.0; TStrong = 0.0; Critical = 0.0
            ScaleGuard = $false; BaselineReliable = $false
        }
    }
    $arr = New-Object double[] $validCount
    [Array]::Copy($valid, $arr, $validCount)
    [Array]::Sort($arr)

    $median = Get-PercentileSorted $arr 0.5
    $dev = New-Object double[] $validCount
    for ($i = 0; $i -lt $validCount; $i++) { $dev[$i] = [Math]::Abs($arr[$i] - $median) }
    [Array]::Sort($dev)
    $mad = Get-PercentileSorted $dev 0.5
    $sigma = 1.4826 * $mad
    $scaleGuard = $false
    if ($sigma -lt [double]$Config.MinRobustSigmaMs) {
        $sigma = [double]$Config.MinRobustSigmaMs
        $scaleGuard = $true
    }
    $p99 = Get-PercentileSorted $arr 0.99
    $p999 = Get-PercentileSorted $arr 0.999
    $tGlobal = [Math]::Max($median + ([double]$Config.GlobalK * $sigma), $p99)
    $tExtremeStat = [Math]::Max($median + ([double]$Config.ExtremeK * $sigma), $p999)
    $critical = [Math]::Max($tExtremeStat, $median * [double]$Config.ExtremeRatioMin)
    $tStrong = [Math]::Max($median * [double]$Config.StrongRatioMin, $tGlobal)
    $reliable = ($validCount -ge 100) -and (($errCount / [double]$n) -lt 0.5)

    return [pscustomobject]@{
        ValidCount = $validCount
        ErrorCount = $errCount
        MissingCount = ($n - $validCount - $errCount)
        Min = [double]$arr[0]
        Max = [double]$arr[$validCount - 1]
        Mean = ($sum / [double]$validCount)
        Median = $median
        MAD = $mad
        RobustSigma = $sigma
        P25 = (Get-PercentileSorted $arr 0.25)
        P50 = $median
        P75 = (Get-PercentileSorted $arr 0.75)
        P90 = (Get-PercentileSorted $arr 0.90)
        P95 = (Get-PercentileSorted $arr 0.95)
        P97 = (Get-PercentileSorted $arr 0.97)
        P98 = (Get-PercentileSorted $arr 0.98)
        P99 = $p99
        P995 = (Get-PercentileSorted $arr 0.995)
        P999 = $p999
        TGlobal = $tGlobal
        TExtremeStat = $tExtremeStat
        TExtreme = $tExtremeStat
        TStrong = $tStrong
        Critical = $critical
        ScaleGuard = $scaleGuard
        BaselineReliable = $reliable
    }
}

# ---------------------------------------------------------------------------
# 加速内核：局部窗口统计（C# 编译；失败时回退纯 PowerShell）
# ---------------------------------------------------------------------------
$script:CoreSource = @'
using System;
public static class DiskGuardCore
{
    public static double Quantile(double[] a, int count, double q)
    {
        if (count <= 0) return 0.0;
        if (count == 1) return a[0];
        double idx = (count - 1) * q;
        int lo = (int)Math.Floor(idx);
        int hi = (int)Math.Ceiling(idx);
        if (lo == hi) return a[lo];
        return a[lo] + (a[hi] - a[lo]) * (idx - lo);
    }

    public static void LocalStats(double[] ms, bool[] valid, int radius, double k, double minSigma, double ratioMin, double minDelta, int minSamples,
        out double[] med, out double[] thr, out double[] p99, out bool[] flag)
    {
        int n = ms.Length;
        med = new double[n];
        thr = new double[n];
        p99 = new double[n];
        flag = new bool[n];
        int cap = 2 * radius + 1;
        double[] buf = new double[cap];
        double[] dev = new double[cap];
        for (int i = 0; i < n; i++)
        {
            if (!valid[i])
            {
                med[i] = double.NaN; thr[i] = double.NaN; p99[i] = double.NaN;
                continue;
            }
            int lo = i - radius; if (lo < 0) lo = 0;
            int hi = i + radius; if (hi > n - 1) hi = n - 1;
            int c = 0;
            for (int j = lo; j <= hi; j++)
            {
                if (j == i || !valid[j]) continue;
                buf[c++] = ms[j];
            }
            if (c < minSamples)
            {
                med[i] = double.NaN; thr[i] = double.NaN; p99[i] = double.NaN;
                continue;
            }
            Array.Sort(buf, 0, c);
            double m = Quantile(buf, c, 0.5);
            for (int j = 0; j < c; j++) dev[j] = Math.Abs(buf[j] - m);
            Array.Sort(dev, 0, c);
            double mad = Quantile(dev, c, 0.5);
            double sigma = Math.Max(1.4826 * mad, minSigma);
            double p = Quantile(buf, c, 0.99);
            double t = Math.Max(m + k * sigma, p);
            med[i] = m; thr[i] = t; p99[i] = p;
            if (ms[i] > t && ms[i] > m * ratioMin && (ms[i] - m) >= minDelta) flag[i] = true;
        }
    }
}
'@

function Initialize-Accelerator {
    try {
        if (-not ('DiskGuardCore' -as [type])) {
            Add-Type -TypeDefinition $script:CoreSource -Language CSharp -ErrorAction Stop
        }
        $script:Accelerated = $true
        Add-Log '加速内核 DiskGuardCore 已加载。'
    } catch {
        $script:Accelerated = $false
        Write-Warn ("无法编译加速内核，改用纯 PowerShell 实现（会明显变慢）：{0}" -f $_.Exception.Message)
    }
}

function Get-LocalStatsSlow {
    param([double[]]$ReadMs, [bool[]]$Valid, [int]$Radius, [double]$K, [double]$MinSigma, [double]$RatioMin, [double]$MinDelta, [int]$MinSamples)
    $n = $ReadMs.Length
    $med = New-Object double[] $n
    $thr = New-Object double[] $n
    $p99 = New-Object double[] $n
    $flag = New-Object bool[] $n
    for ($i = 0; $i -lt $n; $i++) {
        if (-not $Valid[$i]) {
            $med[$i] = [double]::NaN; $thr[$i] = [double]::NaN; $p99[$i] = [double]::NaN
            continue
        }
        $lo = [Math]::Max(0, $i - $Radius)
        $hi = [Math]::Min($n - 1, $i + $Radius)
        $buf = New-Object 'System.Collections.Generic.List[double]'
        for ($j = $lo; $j -le $hi; $j++) {
            if ($j -ne $i -and $Valid[$j]) { $buf.Add($ReadMs[$j]) }
        }
        $c = $buf.Count
        if ($c -lt $MinSamples) {
            $med[$i] = [double]::NaN; $thr[$i] = [double]::NaN; $p99[$i] = [double]::NaN
            continue
        }
        $a = $buf.ToArray()
        [Array]::Sort($a)
        $m = Get-PercentileSorted $a 0.5
        $dev = New-Object double[] $c
        for ($j = 0; $j -lt $c; $j++) { $dev[$j] = [Math]::Abs($a[$j] - $m) }
        [Array]::Sort($dev)
        $mad = Get-PercentileSorted $dev 0.5
        $sigma = [Math]::Max(1.4826 * $mad, $MinSigma)
        $p = Get-PercentileSorted $a 0.99
        $t = [Math]::Max($m + $K * $sigma, $p)
        $med[$i] = $m; $thr[$i] = $t; $p99[$i] = $p
        if ($ReadMs[$i] -gt $t -and $ReadMs[$i] -gt ($m * $RatioMin) -and (($ReadMs[$i] - $m) -ge $MinDelta)) { $flag[$i] = $true }
    }
    return [pscustomobject]@{ Median = $med; Threshold = $thr; P99 = $p99; Flag = $flag }
}

function Get-LocalStats {
    param([double[]]$ReadMs, [bool[]]$Valid, [int]$Radius, [double]$K, [double]$MinSigma, [double]$RatioMin, [double]$MinDelta, [int]$MinSamples)
    if ($script:Accelerated) {
        $med = $null; $thr = $null; $p99 = $null; $flag = $null
        [DiskGuardCore]::LocalStats($ReadMs, $Valid, $Radius, $K, $MinSigma, $RatioMin, $MinDelta, $MinSamples, [ref]$med, [ref]$thr, [ref]$p99, [ref]$flag)
        return [pscustomobject]@{ Median = $med; Threshold = $thr; P99 = $p99; Flag = $flag }
    }
    return (Get-LocalStatsSlow -ReadMs $ReadMs -Valid $Valid -Radius $Radius -K $K -MinSigma $MinSigma -RatioMin $RatioMin -MinDelta $MinDelta -MinSamples $MinSamples)
}

# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------
function Parse-Report {
    param([string]$Path, [string]$Label, $Config)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "找不到报告文件：$Path" }

    $reModel = New-Object System.Text.RegularExpressions.Regex('被检测磁盘\s*[:：]\s*(.+)')
    $reRange = New-Object System.Text.RegularExpressions.Regex('检测范围\s*[:：]\s*柱面\s*(\d+)\s*[-~至]\s*(\d+)')
    $reTime = New-Object System.Text.RegularExpressions.Regex('完成时间\s*[:：]\s*(.+)')
    $reBad = New-Object System.Text.RegularExpressions.Regex('共检测到\s*(\d+)\s*个坏道区域')
    $reNorm = New-Object System.Text.RegularExpressions.Regex('共检测到\s*(\d+)\s*个正常磁道')
    $reRow = New-Object System.Text.RegularExpressions.Regex('^\s*(\d+)\s+(\S+)\s+(\S+)\s+(\d+)\s+(\d+)(?:\s+(\S+)(?:\s+(.*?))?)?\s*$')
    $reRepair = New-Object System.Text.RegularExpressions.Regex('^(未处理|已处理|已修复|未修复|无法修复|修复|成功|N/?A|-+|—)$')
    $reErr = New-Object System.Text.RegularExpressions.Regex('(?i)unknown\s*error|error|错误|失败|坏道|未能读取|读取失败|无法读取|不可读取|超时|crc|坏扇区')

    $enc = Get-TextEncoding $Path
    $cylList = New-Object 'System.Collections.Generic.List[int64]'
    $msList = New-Object 'System.Collections.Generic.List[double]'
    $errList = New-Object 'System.Collections.Generic.List[bool]'
    $txtList = New-Object 'System.Collections.Generic.List[string]'
    $secList = New-Object 'System.Collections.Generic.List[int64]'
    $lineList = New-Object 'System.Collections.Generic.List[int]'
    $repList = New-Object 'System.Collections.Generic.List[string]'
    $unparsed = New-Object 'System.Collections.Generic.List[object]'
    $duplicates = New-Object 'System.Collections.Generic.List[object]'
    $seen = New-Object 'System.Collections.Generic.HashSet[int64]'
    $unparsedCount = 0
    $duplicateCount = 0

    $model = $null; $scanTime = $null
    $rangeStart = $null; $rangeEnd = $null
    $declaredBad = $null; $declaredNormal = $null
    $section = 'header'
    $lineNo = 0
    $errRows = 0

    $sr = New-Object System.IO.StreamReader($Path, $enc, $true)
    try {
        while ($null -ne ($line = $sr.ReadLine())) {
            $lineNo++
            $s = $line.Trim()
            if ($s.Length -eq 0) { continue }

            if ($s.StartsWith('DiskGenius')) { continue }
            if ($s.IndexOf('被检测磁盘', [System.StringComparison]::Ordinal) -ge 0) {
                $m = $reModel.Match($s); if ($m.Success) { $model = $m.Groups[1].Value.Trim() }
                continue
            }
            if ($s.IndexOf('检测范围', [System.StringComparison]::Ordinal) -ge 0) {
                $m = $reRange.Match($s)
                if ($m.Success) { $rangeStart = [int64]$m.Groups[1].Value; $rangeEnd = [int64]$m.Groups[2].Value }
                continue
            }
            if ($s.IndexOf('完成时间', [System.StringComparison]::Ordinal) -ge 0) {
                $m = $reTime.Match($s); if ($m.Success) { $scanTime = $m.Groups[1].Value.Trim() }
                continue
            }
            if ($s.IndexOf('坏道区域', [System.StringComparison]::Ordinal) -ge 0) {
                $m = $reBad.Match($s); if ($m.Success) { $declaredBad = [int64]$m.Groups[1].Value }
                $section = 'bad'
                continue
            }
            if ($s.IndexOf('正常磁道', [System.StringComparison]::Ordinal) -ge 0) {
                $m = $reNorm.Match($s); if ($m.Success) { $declaredNormal = [int64]$m.Groups[1].Value }
                $section = 'normal'
                continue
            }
            if ($s.IndexOf('柱面号', [System.StringComparison]::Ordinal) -ge 0) { continue }

            $m = $reRow.Match($s)
            if (-not $m.Success) {
                $unparsedCount++
                if ($unparsed.Count -lt 50) {
                    $unparsed.Add([pscustomobject]@{ LineNo = $lineNo; Text = $s })
                }
                continue
            }

            $cyl = [int64]$m.Groups[1].Value
            $secCount = [int64]$m.Groups[4].Value
            $msText = $m.Groups[5].Value
            $ms = [double]::NaN
            if (-not [double]::TryParse($msText, [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$ms)) { $ms = [double]::NaN }
            $token = ''
            if ($m.Groups[6].Success) { $token = $m.Groups[6].Value.Trim() }
            $rest = ''
            if ($m.Groups[7].Success) { $rest = $m.Groups[7].Value.Trim() }

            $repair = ''
            $errText = ''
            if ($token.Length -gt 0) {
                if ($reRepair.IsMatch($token)) { $repair = $token; $errText = $rest }
                else { $errText = ($token + ' ' + $rest).Trim() }
            }

            if ($cyl -lt 0 -or $cyl -gt 100000000 -or $secCount -le 0 -or $secCount -gt 10000000 -or [double]::IsNaN($ms) -or $ms -lt 0 -or $ms -gt 86400000) {
                $unparsedCount++
                if ($unparsed.Count -lt 50) { $unparsed.Add([pscustomobject]@{ LineNo = $lineNo; Text = $s }) }
                continue
            }

            if (-not $seen.Add($cyl)) {
                $duplicateCount++
                if ($duplicates.Count -lt 20) { $duplicates.Add([pscustomobject]@{ LineNo = $lineNo; Cylinder = $cyl }) }
                continue
            }

            $isError = ($section -eq 'bad')
            if (-not $isError -and $errText.Length -gt 0 -and $reErr.IsMatch($errText)) { $isError = $true }
            if ($isError) { $errRows++ }

            $cylList.Add($cyl)
            $msList.Add($ms)
            $errList.Add($isError)
            $txtList.Add($errText)
            $secList.Add($secCount)
            $lineList.Add($lineNo)
            $repList.Add($repair)
        }
    } finally {
        $sr.Dispose()
    }

    $count = $cylList.Count
    if ($count -eq 0) { throw "未能从报告中解析出柱面数据。文件：$Path" }

    # 按柱面排序（坏道段通常排在正常段之前，必须重排）。
    $keys = $cylList.ToArray()
    $order = New-Object int[] $count
    for ($i = 0; $i -lt $count; $i++) { $order[$i] = $i }
    [Array]::Sort($keys, $order)

    $cyl = New-Object int64[] $count
    $ms = New-Object double[] $count
    $isErr = New-Object bool[] $count
    $errText = New-Object string[] $count
    $sec = New-Object int64[] $count
    $ln = New-Object int[] $count
    $rep = New-Object string[] $count
    $msRaw = $msList.ToArray()
    $errRaw = $errList.ToArray()
    $txtRaw = $txtList.ToArray()
    $secRaw = $secList.ToArray()
    $lnRaw = $lineList.ToArray()
    $repRaw = $repList.ToArray()
    for ($i = 0; $i -lt $count; $i++) {
        $k = $order[$i]
        $cyl[$i] = $keys[$i]
        $ms[$i] = $msRaw[$k]
        $isErr[$i] = $errRaw[$k]
        $errText[$i] = $txtRaw[$k]
        $sec[$i] = $secRaw[$k]
        $ln[$i] = $lnRaw[$k]
        $rep[$i] = $repRaw[$k]
    }

    # 扇区数众数 = 每柱面扇区数
    $secFreq = New-Object 'System.Collections.Generic.Dictionary[int64,int]'
    for ($i = 0; $i -lt $count; $i++) {
        $v = $sec[$i]
        if ($secFreq.ContainsKey($v)) { $secFreq[$v] = $secFreq[$v] + 1 } else { $secFreq[$v] = 1 }
    }
    $spc = [int64]0; $spcFreq = 0
    foreach ($kv in $secFreq.GetEnumerator()) {
        if ($kv.Value -gt $spcFreq) { $spcFreq = $kv.Value; $spc = $kv.Key }
    }
    $capacity = [int64]0
    for ($i = 0; $i -lt $count; $i++) { $capacity += $sec[$i] }
    $lastSectors = $sec[$count - 1]

    # 连续性与缺失
    $missing = New-Object 'System.Collections.Generic.List[object]'
    $missingCount = 0
    $contiguous = $true
    for ($i = 1; $i -lt $count; $i++) {
        $delta = $cyl[$i] - $cyl[$i - 1]
        if ($delta -ne 1) {
            $contiguous = $false
            $missingCount += [int]($delta - 1)
            if ($missing.Count -lt 20) {
                $missing.Add([pscustomobject]@{ AfterCylinder = $cyl[$i - 1]; BeforeCylinder = $cyl[$i]; Missing = ($delta - 1) })
            }
        }
    }
    $startsAtZero = ($cyl[0] -eq 0)

    $stats = Get-Statistics -ReadMs $ms -IsError $isErr -Config $Config

    $report = [pscustomobject]@{
        Path = $Path
        Label = $Label
        Model = $model
        ScanTime = $scanTime
        RangeStart = $rangeStart
        RangeEnd = $rangeEnd
        DeclaredBad = $declaredBad
        DeclaredNormal = $declaredNormal
        Cylinder = $cyl
        ReadMs = $ms
        IsError = $isErr
        ErrorText = $errText
        SectorCount = $sec
        LineNo = $ln
        RepairStatus = $rep
        Count = $count
        ErrorCount = $errRows
        Unparsed = $unparsed.ToArray()
        UnparsedCount = $unparsedCount
        Duplicates = $duplicates.ToArray()
        DuplicateCount = $duplicateCount
        Missing = $missing.ToArray()
        MissingCount = $missingCount
        Contiguous = $contiguous
        StartsAtZero = $startsAtZero
        SectorsPerCylinder = $spc
        SectorsPerCylinderFreq = $spcFreq
        CapacitySectors = $capacity
        LastCylinderSectors = $lastSectors
        Stats = $stats
        Flags = $null
        Validation = (New-Object 'System.Collections.Generic.List[object]')
        Fatal = $false
    }
    return $report
}

# ---------------------------------------------------------------------------
# 结构校验
# ---------------------------------------------------------------------------
function Test-ReportStructure {
    param($Report, $Config)

    $v = $Report.Validation
    $v.Clear()
    $fatal = $false
    function Add-Check([string]$Level, [string]$Message) {
        $v.Add([pscustomobject]@{ Level = $Level; Message = $Message })
        if ($Level -eq 'error') { $script:StructFatal = $true }
    }
    $script:StructFatal = $false

    if ([string]::IsNullOrWhiteSpace($Report.Model)) {
        Add-Check 'warning' '报告未识别出磁盘型号。'
    }
    if ($null -eq $Report.RangeStart -or $null -eq $Report.RangeEnd) {
        Add-Check 'warning' '报告未识别出扫描范围，LBA 映射无法验证。'
    } else {
        $expected = $Report.RangeEnd - $Report.RangeStart + 1
        if ($expected -ne $Report.Count) {
            Add-Check 'warning' ("扫描范围声明为 {0} 个柱面，实际解析 {1} 条记录。" -f $expected, $Report.Count)
        }
        if (-not $Report.StartsAtZero) {
            Add-Check 'warning' ("扫描未从柱面 0 开始（起始柱面 {0}），不能直接套用 LBA=C*S 映射。" -f $Report.Cylinder[0])
        }
    }
    if ($Report.DuplicateCount -gt 0) {
        Add-Check 'error' ("存在 {0} 条重复柱面记录（例如柱面 {1}），已保留首条。" -f $Report.DuplicateCount, $Report.Duplicates[0].Cylinder)
    }
    if ($Report.MissingCount -gt 0) {
        Add-Check 'warning' ("柱面编号不连续，缺失 {0} 个柱面。" -f $Report.MissingCount)
    }
    if (-not $Report.Contiguous) { Add-Check 'warning' '柱面记录不连续，LBA 映射需要人工确认。' }
    if ($Report.SectorsPerCylinder -le 0) {
        Add-Check 'error' '无法确定每柱面扇区数。'
    }
    if ($Report.LastCylinderSectors -gt $Report.SectorsPerCylinder) {
        Add-Check 'error' '最后一个柱面的扇区数大于常规柱面，记录结构异常。'
    }
    if ($null -ne $Report.DeclaredNormal -and $Report.DeclaredNormal -ne ($Report.Count - $Report.ErrorCount)) {
        Add-Check 'warning' ("报告声明 {0} 条正常磁道，实际解析 {1} 条正常记录。" -f $Report.DeclaredNormal, ($Report.Count - $Report.ErrorCount))
    }
    if ($null -ne $Report.DeclaredBad -and $Report.DeclaredBad -ne $Report.ErrorCount) {
        Add-Check 'warning' ("报告声明 {0} 个坏道区域，实际解析 {1} 条错误记录。" -f $Report.DeclaredBad, $Report.ErrorCount)
    }
    if ($Report.UnparsedCount -gt 0) {
        Add-Check 'warning' ("存在 {0} 行无法解析的记录（已在日志中列出前 {1} 行）。" -f $Report.UnparsedCount, $Report.Unparsed.Count)
    }
    if (-not $Report.Stats.BaselineReliable) {
        Add-Check 'warning' '有效延迟样本不足或错误比例过高，统计基线可信度不足。'
    }
    if ($Report.Stats.ScaleGuard) {
        Add-Check 'info' ("MAD 过小，已启用最小稳健尺度保护（{0} ms）。" -f $Config.MinRobustSigmaMs)
    }
    $Report.Fatal = $script:StructFatal
    return $Report
}

# ---------------------------------------------------------------------------
# 异常检测
# ---------------------------------------------------------------------------
function Add-DetectionFlags {
    param($Report, $Config)

    $n = $Report.Count
    $ms = $Report.ReadMs
    $isErr = $Report.IsError
    $stats = $Report.Stats
    $valid = New-Object bool[] $n
    for ($i = 0; $i -lt $n; $i++) { $valid[$i] = (-not $isErr[$i]) -and (-not [double]::IsNaN($ms[$i])) }

    $small = Get-LocalStats -ReadMs $ms -Valid $valid -Radius ([int]$Config.LocalWindowSmall) -K ([double]$Config.LocalK) -MinSigma ([double]$Config.LocalMinSigmaMs) -RatioMin ([double]$Config.LocalRatioMin) -MinDelta ([double]$Config.LocalMinDeltaMs) -MinSamples ([int]$Config.LocalMinSamples)
    $large = Get-LocalStats -ReadMs $ms -Valid $valid -Radius ([int]$Config.LocalWindowLarge) -K ([double]$Config.LocalK) -MinSigma ([double]$Config.LocalMinSigmaMs) -RatioMin ([double]$Config.LocalRatioMin) -MinDelta ([double]$Config.LocalMinDeltaMs) -MinSamples ([int]$Config.LocalMinSamples)

    $globalFlag = New-Object bool[] $n
    $extremeFlag = New-Object bool[] $n
    $localFlag = New-Object bool[] $n
    $bgFlag = New-Object bool[] $n
    $bgMin = $stats.Median * [double]$Config.RegionSlowRatio
    $elevatedMin = [Math]::Max($stats.Median * [double]$Config.RegionSlowRatio, $stats.Median + (4.0 * $stats.RobustSigma))
    for ($i = 0; $i -lt $n; $i++) {
        if (-not $valid[$i]) { continue }
        $v = $ms[$i]
        $globalFlag[$i] = ($v -ge $stats.TGlobal)
        $extremeFlag[$i] = ($v -ge $stats.TExtreme)
        $localFlag[$i] = ($small.Flag[$i] -or $large.Flag[$i])
        $bgFlag[$i] = (($large.Median[$i] -ge $bgMin) -or ($v -ge $elevatedMin))
    }

    $Report.Flags = [pscustomobject]@{
        Global = $globalFlag
        Extreme = $extremeFlag
        Local = $localFlag
        Background = $bgFlag
        LocalSmall = $small.Flag
        LocalLarge = $large.Flag
        MedianSmall = $small.Median
        ThresholdSmall = $small.Threshold
        MedianLarge = $large.Median
        ThresholdLarge = $large.Threshold
    }
    return $Report
}

# ---------------------------------------------------------------------------
# 双报告对照与异常点集合
# ---------------------------------------------------------------------------
function Get-AnomalySet {
    param($R1, $R2)
    $set = New-Object 'System.Collections.Generic.HashSet[int64]'
    for ($i = 0; $i -lt $R1.Count; $i++) {
        if ($R1.IsError[$i] -or $R1.Flags.Global[$i] -or $R1.Flags.Local[$i]) { [void]$set.Add($R1.Cylinder[$i]) }
    }
    if ($null -ne $R2) {
        for ($i = 0; $i -lt $R2.Count; $i++) {
            if ($R2.IsError[$i] -or $R2.Flags.Global[$i] -or $R2.Flags.Local[$i]) { [void]$set.Add($R2.Cylinder[$i]) }
        }
    }
    $arr = New-Object int64[] $set.Count
    $set.CopyTo($arr)
    [Array]::Sort($arr)
    return ,$arr
}

function Build-Clusters {
    param([int64[]]$Points, [int64]$Gap, [int64]$MaxSpan)
    $clusters = New-Object 'System.Collections.Generic.List[object]'
    if ($null -eq $Points -or $Points.Length -eq 0) { return ,$clusters }
    $start = $Points[0]; $prev = $Points[0]
    $members = New-Object 'System.Collections.Generic.List[int64]'
    $members.Add($Points[0])
    for ($i = 1; $i -lt $Points.Length; $i++) {
        $p = $Points[$i]
        if ((($p - $prev) -gt $Gap) -or (($p - $start) -ge $MaxSpan)) {
            $clusters.Add([pscustomobject]@{ Start = $start; End = $prev; Members = $members.ToArray() })
            $start = $p
            $members = New-Object 'System.Collections.Generic.List[int64]'
        }
        $members.Add($p)
        $prev = $p
    }
    $clusters.Add([pscustomobject]@{ Start = $start; End = $prev; Members = $members.ToArray() })
    return $clusters
}

function Get-ClusterMetrics {
    param([int64]$StartCyl, [int64]$EndCyl, [int64[]]$Members, $R1, $R2, $Config)

    $pad = [int64]$Config.ContextCylinders
    $ctxStart = [Math]::Max([int64]0, $StartCyl - $pad)
    $ctxEnd = [Math]::Min([int64]($R1.Count - 1), $EndCyl + $pad)

    $idx1 = $R1.CylinderIndex
    $idx2 = $null
    if ($null -ne $R2) { $idx2 = $R2.CylinderIndex }

    $vals1 = New-Object 'System.Collections.Generic.List[double]'
    $vals2 = New-Object 'System.Collections.Generic.List[double]'
    $maxMs1 = 0.0; $maxMs2 = 0.0
    $errCount = 0; $extreme1 = $false; $extreme2 = $false
    $crit1 = $false; $crit2 = $false
    $global1 = 0; $global2 = 0; $local1 = 0; $local2 = 0
    $critMs = $R1.Stats.Critical

    foreach ($c in $Members) {
        $key = [string]$c
        $k1 = -1
        if ($idx1.ContainsKey($key)) { $k1 = $idx1[$key] }
        $k2 = -1
        if ($null -ne $R2 -and $idx2.ContainsKey($key)) { $k2 = $idx2[$key] }
        $errHere = $false
        if ($k1 -ge 0) {
            if ($R1.IsError[$k1]) { $errHere = $true }
            else {
                $vals1.Add($R1.ReadMs[$k1])
                if ($R1.ReadMs[$k1] -gt $maxMs1) { $maxMs1 = $R1.ReadMs[$k1] }
                if ($R1.Flags.Extreme[$k1]) { $extreme1 = $true }
                if ($R1.ReadMs[$k1] -ge $critMs) { $crit1 = $true }
                if ($R1.Flags.Global[$k1]) { $global1++ }
                if ($R1.Flags.Local[$k1]) { $local1++ }
            }
        }
        if ($k2 -ge 0) {
            if ($R2.IsError[$k2]) { $errHere = $true }
            else {
                $vals2.Add($R2.ReadMs[$k2])
                if ($R2.ReadMs[$k2] -gt $maxMs2) { $maxMs2 = $R2.ReadMs[$k2] }
                if ($R2.Flags.Extreme[$k2]) { $extreme2 = $true }
                if ($R2.ReadMs[$k2] -ge $R2.Stats.Critical) { $crit2 = $true }
                if ($R2.Flags.Global[$k2]) { $global2++ }
                if ($R2.Flags.Local[$k2]) { $local2++ }
            }
        }
        if ($errHere) { $errCount++ }
    }

    # 区域背景指标（含异常点前后 ContextCylinders 个柱面）
    $region1 = New-Object 'System.Collections.Generic.List[double]'
    $region2 = New-Object 'System.Collections.Generic.List[double]'
    for ($c = $ctxStart; $c -le $ctxEnd; $c++) {
        $key = [string]$c
        if ($idx1.ContainsKey($key)) {
            $k1 = $idx1[$key]
            if (-not $R1.IsError[$k1] -and -not [double]::IsNaN($R1.ReadMs[$k1])) { $region1.Add($R1.ReadMs[$k1]) }
        }
        if ($null -ne $R2 -and $idx2.ContainsKey($key)) {
            $k2 = $idx2[$key]
            if (-not $R2.IsError[$k2] -and -not [double]::IsNaN($R2.ReadMs[$k2])) { $region2.Add($R2.ReadMs[$k2]) }
        }
    }
    $r1arr = $region1.ToArray(); [Array]::Sort($r1arr)
    $r2arr = $region2.ToArray(); [Array]::Sort($r2arr)

    $span = [int64]($EndCyl - $StartCyl + 1)
    $repeat = ($null -ne $R2) -and ($global1 + $local1 + $errCount -gt 0) -and ($global2 + $local2 -gt 0)
    $maxGap = [int64]0
    for ($i = 1; $i -lt $Members.Length; $i++) {
        $g = $Members[$i] - $Members[$i - 1]
        if ($g -gt $maxGap) { $maxGap = $g }
    }

    return [pscustomobject]@{
        Kind = 'Anomaly'
        StartCylinder = $StartCyl
        EndCylinder = $EndCyl
        Span = $span
        PointCount = $Members.Length
        Members = $Members
        ErrorPoints = $errCount
        Extreme1 = $extreme1
        Extreme2 = $extreme2
        Critical1 = $crit1
        Critical2 = $crit2
        Global1 = $global1
        Global2 = $global2
        Local1 = $local1
        Local2 = $local2
        MaxMs1 = $maxMs1
        MaxMs2 = $maxMs2
        RegionMedian1 = (Get-PercentileSorted $r1arr 0.5)
        RegionP951 = (Get-PercentileSorted $r1arr 0.95)
        RegionMedian2 = (Get-PercentileSorted $r2arr 0.5)
        RegionP952 = (Get-PercentileSorted $r2arr 0.95)
        Density = ($Members.Length / [double]$span)
        MaxGap = $maxGap
        Repeat = $repeat
        Risk = 'N'
        RiskReason = ''
        Isolated = $false
        GuardMiB = 0.0
    }
}

function Set-RiskGrades {
    param($Clusters, $R1, $R2, $Config)

    foreach ($cl in $Clusters) {
        $strong = $R1.Stats.TStrong
        if ($null -ne $R2 -and $R2.Stats.TStrong -gt $strong) { $strong = $R2.Stats.TStrong }
        $maxMs = [Math]::Max($cl.MaxMs1, $cl.MaxMs2)
        $hasError = ($cl.ErrorPoints -gt 0)

        if ($hasError) {
            $cl.Risk = 'A'
            $cl.RiskReason = '任意一次扫描存在明确读取错误'
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardErrorMiB
        } elseif ($cl.Critical1 -and $cl.Critical2) {
            $cl.Risk = 'B'
            $cl.RiskReason = '两次扫描均出现极端延迟（区域级重复证据）'
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardExtremeMiB
        } elseif ($cl.Critical1 -or $cl.Critical2) {
            $cl.Risk = 'D'
            $cl.RiskReason = '单次扫描出现极端延迟'
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardExtremeMiB
        } elseif ($cl.PointCount -ge [int]$Config.MinDensePoints -and $maxMs -ge $strong -and ([Math]::Max($cl.RegionP951, $cl.RegionP952) -ge $R1.Stats.TGlobal)) {
            $cl.Risk = 'C'
            $cl.RiskReason = ("异常区域密度较高且幅度显著（最大 {0} ms ≥ {1} ms）" -f (Format-Num $maxMs), (Format-Num $strong))
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardRepeatedMiB
        } elseif ($cl.PointCount -ge [int]$Config.MinDensePoints -or $cl.Extreme1 -or $cl.Extreme2) {
            $cl.Risk = 'E'
            $cl.RiskReason = '统计异常但幅度或密度不足，列入观察'
            $cl.Isolated = $false
            $cl.GuardMiB = 0.0
        } else {
            $cl.Risk = 'F'
            $cl.RiskReason = '孤立单次轻度异常，列入观察'
            $cl.Isolated = $false
            $cl.GuardMiB = 0.0
        }
    }
    return ,$Clusters
}

# ---------------------------------------------------------------------------
# 风险策略：按配置决定哪些等级真正进入隔离
# ---------------------------------------------------------------------------
function Apply-RiskProfile {
    param($Clusters, $Config)
    $profile = [string]$Config.RiskProfile
    $allow = @('A', 'B', 'C', 'D')
    if ($profile -eq 'standard') { $allow = @('A', 'B', 'D') }
    elseif ($profile -eq 'errors-only') { $allow = @('A') }
    foreach ($cl in $Clusters) {
        if ($cl.Isolated -and ($allow -notcontains $cl.Risk)) {
            $cl.Isolated = $false
            $cl.GuardMiB = 0.0
            $cl.RiskReason = $cl.RiskReason + ('；按 RiskProfile={0} 列入观察' -f $profile)
        }
    }
    return ,$Clusters
}

# ---------------------------------------------------------------------------
# 区域整体偏慢（大窗口基线显著高于全盘基线）
# ---------------------------------------------------------------------------
function Get-BackgroundRegions {
    param($R1, $R2, $Config)
    $out = New-Object 'System.Collections.Generic.List[object]'
    $n = $R1.Count
    $bg1 = $R1.Flags.Background
    $gap = [int64]$Config.TightClusterCylinders
    $minSpan = [int64]$Config.MinBackgroundSpan
    $runs = New-Object 'System.Collections.Generic.List[object]'
    $start = $null; $prev = $null
    for ($i = 0; $i -lt $n; $i++) {
        if (-not $bg1[$i]) { continue }
        $c = $R1.Cylinder[$i]
        if ($null -eq $start) { $start = $c; $prev = $c; continue }
        if (($c - $prev) -gt $gap) {
            $runs.Add([pscustomobject]@{ Start = $start; End = $prev })
            $start = $c
        }
        $prev = $c
    }
    if ($null -ne $start) { $runs.Add([pscustomobject]@{ Start = $start; End = $prev }) }

    foreach ($run in $runs) {
        $span = [int64]($run.End - $run.Start + 1)
        if ($span -lt $minSpan) { continue }
        $members = New-Object 'System.Collections.Generic.List[int64]'
        for ($c = $run.Start; $c -le $run.End; $c++) { $members.Add($c) }
        $cl = Get-ClusterMetrics -StartCyl $run.Start -EndCyl $run.End -Members $members.ToArray() -R1 $R1 -R2 $R2 -Config $Config
        $cl.Kind = 'Background'
        $bothScans = $false
        if ($null -ne $R2) {
            $bg2 = $R2.Flags.Background
            $slow = 0; $tot = 0
            for ($c = $run.Start; $c -le $run.End; $c++) {
                $key = [string]$c
                if ($R2.CylinderIndex.ContainsKey($key)) {
                    $tot++
                    if ($bg2[$R2.CylinderIndex[$key]]) { $slow++ }
                }
            }
            if ($tot -gt 0 -and (($slow / [double]$tot) -ge 0.5)) { $bothScans = $true }
        }
        if ($bothScans) {
            $cl.Risk = 'B'
            $cl.RiskReason = '区域整体持续偏慢（连续慢区或大窗口基线偏高），两次扫描均成立'
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardBackgroundMiB
        } else {
            $cl.Risk = 'C'
            $cl.RiskReason = '区域整体偏慢（连续慢区或大窗口基线偏高），仅单次扫描成立'
            $cl.Isolated = $true
            $cl.GuardMiB = [double]$Config.GuardRepeatedMiB
        }
        $out.Add($cl)
    }
    return ,$out
}

# ---------------------------------------------------------------------------
# LBA 映射与隔离区间
# ---------------------------------------------------------------------------
function Convert-LbaToChs {
    param([int64]$Lba, [int64]$Spc, [int]$Heads, [int]$SectorsPerTrack)
    if ($Heads -le 0 -or $SectorsPerTrack -le 0) { return $null }
    if (([int64]$Heads * [int64]$SectorsPerTrack) -ne $Spc) { return $null }
    $c = [int64][Math]::Floor($Lba / [double]$Spc)
    $rem = [int64]($Lba % $Spc)
    $h = [int][Math]::Floor($rem / [double]$SectorsPerTrack)
    $s = [int]($rem % $SectorsPerTrack) + 1
    return [pscustomobject]@{ Cylinder = $c; Head = $h; Sector = $s }
}

function Get-LbaMapping {
    param($R1, $Config)

    $evidence = New-Object 'System.Collections.Generic.List[object]'
    $ok = $true
    function Add-Ev([string]$Name, [bool]$Pass, [string]$Detail) {
        $evidence.Add([pscustomobject]@{ Item = $Name; Pass = $Pass; Detail = $Detail })
    }

    $spc = [int64]$R1.SectorsPerCylinder
    $cap = [int64]$R1.CapacitySectors
    if ($null -ne $Config.SectorsPerCylinderOverride -and [int64]$Config.SectorsPerCylinderOverride -gt 0) {
        $spc = [int64]$Config.SectorsPerCylinderOverride
    }
    if ($null -ne $Config.DiskCapacitySectorsOverride -and [int64]$Config.DiskCapacitySectorsOverride -gt 0) {
        $cap = [int64]$Config.DiskCapacitySectorsOverride
    }

    Add-Ev '扫描从柱面 0 开始' $R1.StartsAtZero ("起始柱面 {0}" -f $R1.Cylinder[0])
    if (-not $R1.StartsAtZero) { $ok = $false }
    Add-Ev '柱面编号连续' $R1.Contiguous ("缺失 {0} 个柱面" -f $R1.MissingCount)
    if (-not $R1.Contiguous) { $ok = $false }
    Add-Ev '每柱面扇区数可确定' ($spc -gt 0) ("S={0}（出现 {1} 次）" -f $spc, $R1.SectorsPerCylinderFreq)
    if ($spc -le 0) { $ok = $false }
    Add-Ev '最后柱面为不完整柱面' (($R1.LastCylinderSectors -gt 0) -and ($R1.LastCylinderSectors -le $spc)) ("最后柱面 {0} 个扇区" -f $R1.LastCylinderSectors)
    if ($R1.LastCylinderSectors -gt $spc) { $ok = $false }
    $calc = [int64]($R1.Count - 1) * $spc + $R1.LastCylinderSectors
    Add-Ev '柱面布局与总扇区数自洽' ($calc -eq $cap) ("(N-1)*S+last={0}，实际合计 {1}" -f $calc, $cap)
    if ($calc -ne $cap) { $ok = $false }
    Add-Ev '逻辑扇区大小有效' ([int]$Config.LogicalSectorBytes -ge 512) ("{0} bytes" -f $Config.LogicalSectorBytes)
    if ([int]$Config.LogicalSectorBytes -lt 512) { $ok = $false }
    if ($R1.DuplicateCount -gt 0) { $ok = $false }
    if ([bool]$Config.RequireManualLbaConfirmation) {
        $manual = [bool]$Config.LbaMappingConfirmed
        Add-Ev '操作者已确认映射' $manual ("LbaMappingConfirmed={0}" -f $manual)
        if (-not $manual) { $ok = $false }
    } else {
        Add-Ev '自动验证模式' $true 'RequireManualLbaConfirmation=false；以上结构证据全部通过即视为已验证'
    }

    return [pscustomobject]@{
        Verified = $ok
        SectorsPerCylinder = $spc
        CapacitySectors = $cap
        Evidence = $evidence.ToArray()
    }
}

function Build-IsolationRanges {
    param($Clusters, $Mapping, $Config, $LastCylinder)

    $spc = [int64]$Mapping.SectorsPerCylinder
    $logical = [int64]$Config.LogicalSectorBytes
    $raw = New-Object 'System.Collections.Generic.List[object]'
    $id = 0
    $ci = 0

    foreach ($cl in $Clusters) {
        $ci++
        if (-not $cl.Isolated) { continue }
        $id++
        $guardSectors = [int64][Math]::Ceiling(($cl.GuardMiB * $script:MiB) / [double]$logical)
        $guardCyl = [int64][Math]::Ceiling($guardSectors / [double]$spc)
        $s = [Math]::Max([int64]0, $cl.StartCylinder - $guardCyl)
        $e = [Math]::Min([int64]$LastCylinder, $cl.EndCylinder + $guardCyl)
        $raw.Add([pscustomobject]@{
            RangeId = $id
            ClusterIds = @($ci)
            StartCylinder = [int64]$s
            EndCylinder = [int64]$e
            RawStart = $cl.StartCylinder
            RawEnd = $cl.EndCylinder
            Risk = $cl.Risk
            GuardMiB = $cl.GuardMiB
            Reason = $cl.RiskReason
            MergedFrom = @($cl.Risk)
            IsMerged = $false
            MaxMs1 = [double]$cl.MaxMs1
            MaxMs2 = [double]$cl.MaxMs2
            RegionP95 = [double]([Math]::Max($cl.RegionP951, $cl.RegionP952))
            Points = [int]$cl.PointCount
            ErrorPoints = [int]$cl.ErrorPoints
            Kind = [string]$cl.Kind
            Members = @($cl.Members)
        })
    }

    # 合并重叠或相接的隔离区间（保留最高风险等级）
    $rank = @{ 'A' = 5; 'B' = 4; 'D' = 3; 'C' = 2 }
    $merged = New-Object 'System.Collections.Generic.List[object]'
    foreach ($rg in @($raw | Sort-Object StartCylinder)) {
        if ($merged.Count -gt 0 -and $rg.StartCylinder -le ($merged[$merged.Count - 1].EndCylinder + 1)) {
            $last = $merged[$merged.Count - 1]
            if ($rg.EndCylinder -gt $last.EndCylinder) { $last.EndCylinder = $rg.EndCylinder }
            $last.ClusterIds = @($last.ClusterIds + $rg.ClusterIds)
            $last.MergedFrom = @($last.MergedFrom + $rg.MergedFrom)
            if ($rg.GuardMiB -gt $last.GuardMiB) { $last.GuardMiB = $rg.GuardMiB }
            if ($rg.MaxMs1 -gt $last.MaxMs1) { $last.MaxMs1 = $rg.MaxMs1 }
            if ($rg.MaxMs2 -gt $last.MaxMs2) { $last.MaxMs2 = $rg.MaxMs2 }
            if ($rg.RegionP95 -gt $last.RegionP95) { $last.RegionP95 = $rg.RegionP95 }
            $last.Points = [int]$last.Points + [int]$rg.Points
            $last.ErrorPoints = [int]$last.ErrorPoints + [int]$rg.ErrorPoints
            if ($rg.Kind -eq 'Background') { $last.Kind = 'Background' }
            $last.Members = @(@($last.Members) + @($rg.Members) | Sort-Object -Unique)
            if ($rank[$rg.Risk] -gt $rank[$last.Risk]) { $last.Risk = $rg.Risk; $last.Reason = $rg.Reason }
            $last.IsMerged = $true
        } else {
            $merged.Add($rg)
        }
    }
    foreach ($m in $merged) { if ($m.IsMerged) { $m.Reason = '合并多个隔离区间；' + $m.Reason } }

    # 计算 LBA 边界（仅在映射已验证时）
    $alignment = [int64][Math]::Ceiling($script:MiB / [double]$logical)
    $firstUsable = [int64]([Math]::Ceiling(34.0 / $alignment) * $alignment)
    $lastUsable = [int64]([Math]::Floor(($Mapping.CapacitySectors - 34) / [double]$alignment) * $alignment - 1)
    $out = New-Object 'System.Collections.Generic.List[object]'
    $rid = 0
    foreach ($rg in $merged) {
        $rid++
        $lbaStart = $null; $lbaEnd = $null; $lbaSectors = $null; $sizeMiB = $null; $status = '仅柱面范围'
        if ($Mapping.Verified) {
            $rs = [int64]$rg.StartCylinder * $spc
            $re = [Math]::Min([int64]($Mapping.CapacitySectors - 1), ([int64]$rg.EndCylinder + 1) * $spc - 1)
            $rs = [Math]::Max($firstUsable, [int64]([Math]::Floor($rs / [double]$alignment) * $alignment))
            $re = [Math]::Min($lastUsable, [int64]([Math]::Ceiling(($re + 1) / [double]$alignment) * $alignment - 1))
            if ($re -ge $rs) {
                $lbaStart = $rs; $lbaEnd = $re; $lbaSectors = $re - $rs + 1
                $sizeMiB = ($lbaSectors * $logical) / $script:MiB
                $status = '按 1 MiB 向外对齐'
            } else {
                $status = '区间落在 GPT 保留区内，已忽略'
            }
        }
        $out.Add([pscustomobject]@{
            RangeId = $rid
            StartCylinder = $rg.StartCylinder
            EndCylinder = $rg.EndCylinder
            Risk = $rg.Risk
            GuardMiBEachSide = $rg.GuardMiB
            Reason = $rg.Reason
            ClusterIds = ($rg.ClusterIds -join ';')
            MergedFrom = ($rg.MergedFrom -join ';')
            LbaStart = $lbaStart
            LbaEnd = $lbaEnd
            LbaSectors = $lbaSectors
            SizeMiB = $sizeMiB
            Status = $status
            MaxMs1 = $rg.MaxMs1
            MaxMs2 = $rg.MaxMs2
            RegionP95 = $rg.RegionP95
            Points = $rg.Points
            ErrorPoints = $rg.ErrorPoints
            Kind = $rg.Kind
            Members = @($rg.Members)
        })
    }
    return [pscustomobject]@{
        Ranges = $out
        FirstUsable = $firstUsable
        LastUsable = $lastUsable
        Alignment = $alignment
    }
}

# ---------------------------------------------------------------------------
# CSV / 文本 / HTML 输出
# ---------------------------------------------------------------------------
function Write-CsvFile {
    param([string]$Path, [string[]]$Columns, $Rows)
    $enc = New-Object System.Text.UTF8Encoding($true)
    $sw = New-Object System.IO.StreamWriter($Path, $false, $enc)
    try {
        $sw.WriteLine(($Columns | ForEach-Object { ConvertTo-CsvText $_ }) -join ',')
        foreach ($row in $Rows) {
            $vals = New-Object string[] $Columns.Length
            for ($i = 0; $i -lt $Columns.Length; $i++) {
                $p = $row.PSObject.Properties[$Columns[$i]]
                if ($null -eq $p) { $vals[$i] = '' } else { $vals[$i] = ConvertTo-CsvText $p.Value }
            }
            $sw.WriteLine(($vals -join ','))
        }
    } finally {
        $sw.Dispose()
    }
}

function Write-CylinderComparison {
    param([string]$Path, $R1, $R2)
    $enc = New-Object System.Text.UTF8Encoding($true)
    $sw = New-Object System.IO.StreamWriter($Path, $false, $enc)
    $fields = New-Object string[] 16
    try {
        $sw.WriteLine('Cylinder,SectorCount,Scan1Ms,Scan1Error,Scan1Global,Scan1Extreme,Scan1Local,Scan1LocalMedian,Scan1LocalThreshold,Scan2Ms,Scan2Error,Scan2Global,Scan2Extreme,Scan2Local,Scan2LocalMedian,Scan2LocalThreshold')
        $idx2 = $null
        if ($null -ne $R2) { $idx2 = $R2.CylinderIndex }
        $n = $R1.Count
        $bt = @('false', 'true')
        $cyl1 = $R1.Cylinder; $sec1 = $R1.SectorCount; $ms1 = $R1.ReadMs; $err1 = $R1.IsError
        $f1 = $R1.Flags
        $g1 = $f1.Global; $x1 = $f1.Extreme; $l1 = $f1.Local; $m1 = $f1.MedianSmall; $t1 = $f1.ThresholdSmall
        $cyl2 = $null; $ms2 = $null; $err2 = $null; $g2 = $null; $x2 = $null; $l2 = $null; $m2 = $null; $t2 = $null
        if ($null -ne $R2) {
            $cyl2 = $R2.Cylinder; $ms2 = $R2.ReadMs; $err2 = $R2.IsError
            $f2 = $R2.Flags
            $g2 = $f2.Global; $x2 = $f2.Extreme; $l2 = $f2.Local; $m2 = $f2.MedianSmall; $t2 = $f2.ThresholdSmall
        }
        for ($i = 0; $i -lt $n; $i++) {
            $fields[0] = $cyl1[$i].ToString()
            $fields[1] = $sec1[$i].ToString()
            if ($err1[$i]) {
                $fields[2] = ''; $fields[3] = 'true'; $fields[4] = 'false'; $fields[5] = 'false'; $fields[6] = 'false'
                $fields[7] = ''; $fields[8] = ''
            } else {
                $fields[2] = $ms1[$i].ToString('0.###', $script:Inv)
                $fields[3] = 'false'
                $fields[4] = $bt[[int]$g1[$i]]
                $fields[5] = $bt[[int]$x1[$i]]
                $fields[6] = $bt[[int]$l1[$i]]
                $fields[7] = $m1[$i].ToString('0.###', $script:Inv)
                $fields[8] = $t1[$i].ToString('0.###', $script:Inv)
            }
            if ($null -eq $idx2) {
                $fields[9] = ''; $fields[10] = ''; $fields[11] = ''; $fields[12] = ''; $fields[13] = ''; $fields[14] = ''; $fields[15] = ''
            } else {
                $k = -1
                if ($idx2.ContainsKey($cyl1[$i].ToString())) { $k = $idx2[$cyl1[$i].ToString()] }
                if ($k -lt 0) {
                    $fields[9] = ''; $fields[10] = ''; $fields[11] = ''; $fields[12] = ''; $fields[13] = ''; $fields[14] = ''; $fields[15] = ''
                } elseif ($err2[$k]) {
                    $fields[9] = ''; $fields[10] = 'true'; $fields[11] = 'false'; $fields[12] = 'false'; $fields[13] = 'false'
                    $fields[14] = ''; $fields[15] = ''
                } else {
                    $fields[9] = $ms2[$k].ToString('0.###', $script:Inv)
                    $fields[10] = 'false'
                    $fields[11] = $bt[[int]$g2[$k]]
                    $fields[12] = $bt[[int]$x2[$k]]
                    $fields[13] = $bt[[int]$l2[$k]]
                    $fields[14] = $m2[$k].ToString('0.###', $script:Inv)
                    $fields[15] = $t2[$k].ToString('0.###', $script:Inv)
                }
            }
            $sw.WriteLine([string]::Join(',', $fields))
        }
    } finally {
        $sw.Dispose()
    }
}

# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
function Invoke-DiskGuard {
    param(
        [string]$Report1,
        [string]$Report2,
        [string]$OutputRoot,
        [string]$ConfigPath,
        [switch]$NoBrowser,
        [switch]$SingleScan,
        [switch]$NoAccelerator
    )

    $script:RunClock.Reset(); $script:RunClock.Start()
    $script:LogLines.Clear()
    $base = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($base)) { $base = (Get-Location).Path }
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $base 'Config.json' }
    $config = Get-DiskGuardConfig $ConfigPath

    if (-not $Report1) { $Report1 = Select-Report '选择第一次扫描报告（必选）' $true }
    elseif (-not (Test-Path -LiteralPath $Report1 -PathType Leaf)) { throw "找不到第一次扫描报告：$Report1" }
    if (-not $Report2 -and -not $SingleScan) { $Report2 = Select-Report '选择第二次扫描报告（可选，取消则单次分析）' $false }
    if (-not $OutputRoot) { $OutputRoot = Join-Path $base 'Output' }
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:FailureLog = Join-Path $OutputRoot ("DiskGuard_failure_{0}.log" -f $stamp)
    [System.IO.File]::WriteAllLines($script:FailureLog, @(
            ('DiskGuard v{0} startup: {1}' -f $script:Version, (Get-Date -Format o)),
            ('Report1: {0}' -f $Report1),
            ('Report2: {0}' -f $Report2)
        ), [System.Text.Encoding]::UTF8)

    if (-not $NoAccelerator) { Initialize-Accelerator }
    else { Write-Warn '已禁用加速内核，将使用纯 PowerShell 局部统计（较慢）。' }

    Write-Info ('DiskGuard v{0} 启动；模式：{1}' -f $script:Version, $(if ($Report2) { '双报告' } else { '单报告' }))

    # ---- 阶段 1：解析 ----
    Write-Info '阶段 1/6：解析报告与结构校验……'
    $script:StageClock.Reset(); $script:StageClock.Start()
    $r1 = Parse-Report -Path $Report1 -Label '扫描1' -Config $config
    Write-Info ('扫描1 解析完成：{0} 条记录（错误 {1} 条），耗时 {2:N2}s' -f $r1.Count, $r1.ErrorCount, $script:StageClock.Elapsed.TotalSeconds)
    $r1 = Test-ReportStructure -Report $r1 -Config $config
    foreach ($item in $r1.Validation) {
        Add-Log ('校验[' + $item.Level + '] ' + $item.Message)
        if ($item.Level -eq 'error') { Write-Warn ('结构校验：' + $item.Message) }
    }

    $r2 = $null
    if ($Report2) {
        $script:StageClock.Restart()
        $r2 = Parse-Report -Path $Report2 -Label '扫描2' -Config $config
        Write-Info ('扫描2 解析完成：{0} 条记录（错误 {1} 条），耗时 {2:N2}s' -f $r2.Count, $r2.ErrorCount, $script:StageClock.Elapsed.TotalSeconds)
        $r2 = Test-ReportStructure -Report $r2 -Config $config
    }

    # 双报告一致性
    $comparisonValid = $true
    $compareNotes = New-Object 'System.Collections.Generic.List[string]'
    if ($null -ne $r2) {
        if ($r1.Model -ne $r2.Model) { $comparisonValid = $false; $compareNotes.Add('两次报告的磁盘型号不一致。') }
        if ($r1.SectorsPerCylinder -ne $r2.SectorsPerCylinder) { $comparisonValid = $false; $compareNotes.Add('两次报告的每柱面扇区数不一致。') }
        if ($r1.RangeStart -ne $r2.RangeStart -or $r1.RangeEnd -ne $r2.RangeEnd) { $comparisonValid = $false; $compareNotes.Add('两次报告的扫描范围不一致。') }
        if ($r1.Count -ne $r2.Count) { $compareNotes.Add('两次报告记录数不同，需人工核查。') }
        foreach ($n in $compareNotes) { Write-Warn ('双报告校验：' + $n) }
    }

    # ---- 阶段 2：全盘统计 ----
    Write-Info '阶段 2/6：建立全盘统计基线……'
    $s1 = $r1.Stats
    Write-Info ('扫描1 基线：中位数 {0} ms，MAD {1} ms，P99 {2} ms，全盘阈值 {3} ms，极端阈值 {4} ms' -f (Format-Num $s1.Median), (Format-Num $s1.MAD), (Format-Num $s1.P99), (Format-Num $s1.TGlobal), (Format-Num $s1.TExtreme))

    # ---- 阶段 3：异常检测 ----
    Write-Info '阶段 3/6：全盘异常与局部异常检测……'
    $script:StageClock.Restart()
    $r1 = Add-DetectionFlags -Report $r1 -Config $config
    if ($null -ne $r2) { $r2 = Add-DetectionFlags -Report $r2 -Config $config }
    Write-Info ('异常检测完成，耗时 {0:N2}s' -f $script:StageClock.Elapsed.TotalSeconds)

    # 柱面索引
    $idx1 = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    for ($i = 0; $i -lt $r1.Count; $i++) { $idx1[[string]$r1.Cylinder[$i]] = $i }
    $r1 | Add-Member -NotePropertyName CylinderIndex -NotePropertyValue $idx1 -Force
    if ($null -ne $r2) {
        $idx2 = New-Object 'System.Collections.Generic.Dictionary[string,int]'
        for ($i = 0; $i -lt $r2.Count; $i++) { $idx2[[string]$r2.Cylinder[$i]] = $i }
        $r2 | Add-Member -NotePropertyName CylinderIndex -NotePropertyValue $idx2 -Force
    }

    # ---- 阶段 4：对照、区域分析、聚类 ----
    Write-Info '阶段 4/6：双报告对照、区域分析与一维聚类……'
    $script:StageClock.Restart()
    $points = Get-AnomalySet -R1 $r1 -R2 $r2
    $baseClusters = Build-Clusters -Points $points -Gap ([int64]$config.TightClusterCylinders) -MaxSpan ([int64]$config.MaxClusterSpan)
    $clusters = New-Object 'System.Collections.Generic.List[object]'
    foreach ($bc in $baseClusters) {
        $clusters.Add((Get-ClusterMetrics -StartCyl $bc.Start -EndCyl $bc.End -Members $bc.Members -R1 $r1 -R2 $r2 -Config $config))
    }
    $clusters = Set-RiskGrades -Clusters $clusters -R1 $r1 -R2 $r2 -Config $config
    $bgClusters = Get-BackgroundRegions -R1 $r1 -R2 $r2 -Config $config
    foreach ($bc in $bgClusters) { $clusters.Add($bc) }
    $clusters = @($clusters | Sort-Object StartCylinder)
    $clusters = Apply-RiskProfile -Clusters $clusters -Config $config
    Write-Info ('聚类完成：{0} 个异常簇（A {1} / B {2} / C {3} / D {4} / E {5} / F {6}），耗时 {7:N2}s' -f `
            $clusters.Count,
        @($clusters | Where-Object Risk -eq 'A').Count,
        @($clusters | Where-Object Risk -eq 'B').Count,
        @($clusters | Where-Object Risk -eq 'C').Count,
        @($clusters | Where-Object Risk -eq 'D').Count,
        @($clusters | Where-Object Risk -eq 'E').Count,
        @($clusters | Where-Object Risk -eq 'F').Count,
        $script:StageClock.Elapsed.TotalSeconds)

    # ---- 阶段 5：隔离区间与 LBA ----
    Write-Info '阶段 5/6：生成安全隔离区间与 LBA 映射……'
    $script:StageClock.Restart()
    $mapping = Get-LbaMapping -R1 $r1 -Config $config
    if (-not $comparisonValid) { $mapping.Verified = $false }
    $lastCyl = $r1.Cylinder[$r1.Count - 1]
    $iso = Build-IsolationRanges -Clusters $clusters -Mapping $mapping -Config $config -LastCylinder $lastCyl
    Write-Info ('隔离区间 {0} 个；LBA 映射：{1}；耗时 {2:N2}s' -f $iso.Ranges.Count, $(if ($mapping.Verified) { '已验证' } else { '未验证' }), $script:StageClock.Elapsed.TotalSeconds)

    # 隔离区异常柱面明细（柱面号 + 两次扫描值）
    $isoDetails = New-Object 'System.Collections.Generic.List[object]'
    $crit1 = [double]$r1.Stats.Critical
    $crit2 = if ($null -ne $r2) { [double]$r2.Stats.Critical } else { [double]::MaxValue }
    foreach ($rg in $iso.Ranges) {
        foreach ($c in @($rg.Members | Sort-Object)) {
            $key = [string]$c
            $k1 = -1
            if ($r1.CylinderIndex.ContainsKey($key)) { $k1 = $r1.CylinderIndex[$key] }
            $k2 = -1
            if ($null -ne $r2 -and $r2.CylinderIndex.ContainsKey($key)) { $k2 = $r2.CylinderIndex[$key] }
            $ms1 = $null; $ms2 = $null; $e1 = $false; $e2 = $false; $x1 = $false; $x2 = $false
            if ($k1 -ge 0) {
                $e1 = $r1.IsError[$k1]
                $ms1 = $r1.ReadMs[$k1]
                if (-not $e1) { $x1 = $r1.Flags.Extreme[$k1] }
            }
            if ($k2 -ge 0) {
                $e2 = $r2.IsError[$k2]
                $ms2 = $r2.ReadMs[$k2]
                if (-not $e2) { $x2 = $r2.Flags.Extreme[$k2] }
            }
            $isCrit = (($null -ne $ms1) -and ($ms1 -ge $crit1)) -or (($null -ne $ms2) -and ($ms2 -ge $crit2))
            $cls = '偏慢'
            if ($e1 -or $e2) { $cls = '读取错误' }
            elseif ($isCrit) { $cls = '严重' }
            elseif ($x1 -or $x2) { $cls = '很慢' }
            $isoDetails.Add([pscustomobject]@{
                    RangeId = $rg.RangeId; RangeStart = $rg.StartCylinder; RangeEnd = $rg.EndCylinder; Risk = $rg.Risk
                    Cylinder = [int64]$c; Scan1Ms = $ms1; Scan2Ms = $ms2
                    Error1 = $e1; Error2 = $e2; Class = $cls
                })
        }
    }

    # 可用分区候选（含起始柱面/磁头/扇区与传统几何）
    $extents = New-Object 'System.Collections.Generic.List[object]'
    $partitions = New-Object 'System.Collections.Generic.List[object]'
    if ($mapping.Verified) {
        $cursor = $iso.FirstUsable
        $extNo = 0
        foreach ($rg in @($iso.Ranges | Where-Object { $null -ne $_.LbaStart } | Sort-Object LbaStart)) {
            if ($rg.LbaStart -gt $cursor) {
                $extNo++
                $extents.Add([pscustomobject]@{ ExtentId = $extNo; LbaStart = $cursor; LbaEnd = ($rg.LbaStart - 1) })
            }
            if (($rg.LbaEnd + 1) -gt $cursor) { $cursor = $rg.LbaEnd + 1 }
        }
        if ($cursor -le $iso.LastUsable) {
            $extNo++
            $extents.Add([pscustomobject]@{ ExtentId = $extNo; LbaStart = $cursor; LbaEnd = $iso.LastUsable })
        }
        $totalGiB = ($mapping.CapacitySectors * [double]$config.LogicalSectorBytes) / 1073741824.0
        if ($totalGiB -le 0) { $totalGiB = 1.0 }
        $partNo = 0
        foreach ($e in $extents) {
            $sectors = [int64]($e.LbaEnd - $e.LbaStart + 1)
            $sizeGiB = ($sectors * [double]$config.LogicalSectorBytes) / 1073741824.0
            $chsS = Convert-LbaToChs -Lba $e.LbaStart -Spc ([int64]$mapping.SectorsPerCylinder) -Heads ([int]$config.HeadsPerCylinder) -SectorsPerTrack ([int]$config.SectorsPerTrack)
            $chsE = Convert-LbaToChs -Lba $e.LbaEnd -Spc ([int64]$mapping.SectorsPerCylinder) -Heads ([int]$config.HeadsPerCylinder) -SectorsPerTrack ([int]$config.SectorsPerTrack)
            $recommended = ($sizeGiB -ge [double]$config.MinPartitionGiB)
            $percent = ($sizeGiB / $totalGiB) * 100.0
            $displayed = ($recommended -and ($percent -ge [double]$config.MinPartitionPercent))
            if ($displayed) { $partNo++ }
            $partitions.Add([pscustomobject]@{
                    Partition = $(if ($displayed) { $partNo } else { 0 })
                    Recommended = $recommended
                    Displayed = $displayed
                    StartCylinder = $(if ($null -ne $chsS) { $chsS.Cylinder } else { $null })
                    StartHead = $(if ($null -ne $chsS) { $chsS.Head } else { $null })
                    StartSector = $(if ($null -ne $chsS) { $chsS.Sector } else { $null })
                    StartLba = [int64]$e.LbaStart
                    TotalSectors = $sectors
                    EndCylinder = $(if ($null -ne $chsE) { $chsE.Cylinder } else { $null })
                    EndHead = $(if ($null -ne $chsE) { $chsE.Head } else { $null })
                    EndSector = $(if ($null -ne $chsE) { $chsE.Sector } else { $null })
                    EndLba = [int64]$e.LbaEnd
                    SizeGiB = $sizeGiB
                    Percent = $percent
                })
        }
    }

    # ---- 阶段 6：输出 ----
    Write-Info '阶段 6/6：写出分析文件……'
    $script:StageClock.Restart()
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeModel = if ([string]::IsNullOrWhiteSpace($r1.Model)) { 'UnknownDisk' } else { ($r1.Model -replace '[^\w\-]', '_') }
    $out = Join-Path $OutputRoot ("{0}_{1}" -f $safeModel, $stamp)
    New-Item -ItemType Directory -Path $out -Force | Out-Null

    $scanRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($r in @($r1, $r2)) {
        if ($null -eq $r) { continue }
        $st = $r.Stats
        $scanRows.Add([pscustomobject]@{
                Scan = $r.Label; Model = $r.Model; ScanTime = $r.ScanTime
                RangeStart = $r.RangeStart; RangeEnd = $r.RangeEnd
                Records = $r.Count; ErrorRecords = $r.ErrorCount; ValidRecords = $st.ValidCount
                MinMs = $st.Min; MeanMs = $st.Mean; MedianMs = $st.Median; MAD = $st.MAD; RobustSigma = $st.RobustSigma
                P25 = $st.P25; P75 = $st.P75; P90 = $st.P90; P95 = $st.P95; P97 = $st.P97; P98 = $st.P98
                P99 = $st.P99; P995 = $st.P995; P999 = $st.P999; MaxMs = $st.Max
                TGlobal = $st.TGlobal; TExtreme = $st.TExtreme; Critical = $st.Critical; TStrong = $st.TStrong
                ScaleGuard = $st.ScaleGuard; BaselineReliable = $st.BaselineReliable
            })
    }
    Write-CsvFile -Path (Join-Path $out 'Scan_Comparison.csv') -Columns @('Scan', 'Model', 'ScanTime', 'RangeStart', 'RangeEnd', 'Records', 'ErrorRecords', 'ValidRecords', 'MinMs', 'MeanMs', 'MedianMs', 'MAD', 'RobustSigma', 'P25', 'P75', 'P90', 'P95', 'P97', 'P98', 'P99', 'P995', 'P999', 'MaxMs', 'TGlobal', 'TExtreme', 'Critical', 'TStrong', 'ScaleGuard', 'BaselineReliable') -Rows $scanRows

    # 异常明细
    $detailRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($c in $points) {
        $key = [string]$c
        $k1 = -1
        if ($r1.CylinderIndex.ContainsKey($key)) { $k1 = $r1.CylinderIndex[$key] }
        $k2 = -1
        if ($null -ne $r2 -and $r2.CylinderIndex.ContainsKey($key)) { $k2 = $r2.CylinderIndex[$key] }
        $ms1 = $null; $ms2 = $null; $e1 = $false; $e2 = $false
        $g1 = $false; $g2 = $false; $x1 = $false; $x2 = $false; $l1 = $false; $l2 = $false
        $lm1 = $null; $lt1 = $null; $lm2 = $null; $lt2 = $null
        if ($k1 -ge 0) {
            $e1 = $r1.IsError[$k1]; $g1 = $r1.Flags.Global[$k1]; $x1 = $r1.Flags.Extreme[$k1]; $l1 = $r1.Flags.Local[$k1]
            if (-not $e1) { $ms1 = $r1.ReadMs[$k1]; $lm1 = $r1.Flags.MedianSmall[$k1]; $lt1 = $r1.Flags.ThresholdSmall[$k1] }
        }
        if ($k2 -ge 0) {
            $e2 = $r2.IsError[$k2]; $g2 = $r2.Flags.Global[$k2]; $x2 = $r2.Flags.Extreme[$k2]; $l2 = $r2.Flags.Local[$k2]
            if (-not $e2) { $ms2 = $r2.ReadMs[$k2]; $lm2 = $r2.Flags.MedianSmall[$k2]; $lt2 = $r2.Flags.ThresholdSmall[$k2] }
        }
        $cls = '异常'
        if ($e1 -or $e2) { $cls = '明确读取错误' }
        elseif ($x1 -and $x2) { $cls = '两次极端' }
        elseif ($x1 -or $x2) { $cls = '单次极端' }
        elseif ($g1 -and $g2) { $cls = '两次全盘异常' }
        elseif ($g1 -or $g2) { $cls = '单次全盘异常' }
        elseif ($l1 -and $l2) { $cls = '两次局部异常' }
        else { $cls = '局部异常' }
        $detailRows.Add([pscustomobject]@{
                Cylinder = $c; Class = $cls
                Scan1Ms = $ms1; Scan2Ms = $ms2
                Error1 = $e1; Error2 = $e2
                Global1 = $g1; Global2 = $g2
                Extreme1 = $x1; Extreme2 = $x2
                Local1 = $l1; Local2 = $l2
                LocalMedian1 = $lm1; LocalThreshold1 = $lt1
                LocalMedian2 = $lm2; LocalThreshold2 = $lt2
            })
    }
    Write-CsvFile -Path (Join-Path $out 'Anomaly_Details.csv') -Columns @('Cylinder', 'Class', 'Scan1Ms', 'Scan2Ms', 'Error1', 'Error2', 'Global1', 'Global2', 'Extreme1', 'Extreme2', 'Local1', 'Local2', 'LocalMedian1', 'LocalThreshold1', 'LocalMedian2', 'LocalThreshold2') -Rows $detailRows

    $clusterRows = New-Object 'System.Collections.Generic.List[object]'
    $cid = 0
    foreach ($cl in $clusters) {
        $cid++
        $clusterRows.Add([pscustomobject]@{
                ClusterId = $cid; Kind = $cl.Kind; StartCylinder = $cl.StartCylinder; EndCylinder = $cl.EndCylinder; Span = $cl.Span
                AnomalyPoints = $cl.PointCount; ErrorPoints = $cl.ErrorPoints
                MaxMs1 = $cl.MaxMs1; MaxMs2 = $cl.MaxMs2
                RegionMedian1 = $cl.RegionMedian1; RegionP951 = $cl.RegionP951
                RegionMedian2 = $cl.RegionMedian2; RegionP952 = $cl.RegionP952
                Density = $cl.Density; MaxGap = $cl.MaxGap; Repeat = $cl.Repeat
                Risk = $cl.Risk; RiskReason = $cl.RiskReason; Isolated = $cl.Isolated
            })
    }
    Write-CsvFile -Path (Join-Path $out 'Anomaly_Clusters.csv') -Columns @('ClusterId', 'Kind', 'StartCylinder', 'EndCylinder', 'Span', 'AnomalyPoints', 'ErrorPoints', 'MaxMs1', 'MaxMs2', 'RegionMedian1', 'RegionP951', 'RegionMedian2', 'RegionP952', 'Density', 'MaxGap', 'Repeat', 'Risk', 'RiskReason', 'Isolated') -Rows $clusterRows

    Write-CsvFile -Path (Join-Path $out 'Isolation_Ranges.csv') -Columns @('RangeId', 'StartCylinder', 'EndCylinder', 'Risk', 'MaxMs1', 'MaxMs2', 'Points', 'RegionP95', 'ErrorPoints', 'Kind', 'GuardMiBEachSide', 'Reason', 'ClusterIds', 'MergedFrom', 'LbaStart', 'LbaEnd', 'LbaSectors', 'SizeMiB', 'Status') -Rows $iso.Ranges
    Write-CsvFile -Path (Join-Path $out 'Isolation_Anomalies.csv') -Columns @('RangeId', 'RangeStart', 'RangeEnd', 'Risk', 'Cylinder', 'Scan1Ms', 'Scan2Ms', 'Error1', 'Error2', 'Class') -Rows $isoDetails

    if ([bool]$config.WriteCylinderComparison) {
        Write-CylinderComparison -Path (Join-Path $out 'Cylinder_Comparison.csv') -R1 $r1 -R2 $r2
    }

# ---- 分区规划（简洁版） ----
    $errList = @()
    for ($i = 0; $i -lt $r1.Count; $i++) { if ($r1.IsError[$i]) { $errList += [pscustomobject]@{ Cylinder = $r1.Cylinder[$i]; Ms = $r1.ReadMs[$i]; Text = $r1.ErrorText[$i] } } }

    $plan = New-Object 'System.Collections.Generic.List[string]'
    $plan.Add(('DiskGuard v{0} 分区方案（只读分析，不会修改磁盘）' -f $script:Version))
    $plan.Add(('生成时间：{0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    $plan.Add(('报告1：{0}' -f $Report1))
    if ($Report2) { $plan.Add(('报告2：{0}' -f $Report2)) }
    $plan.Add('')
    $plan.Add('一、磁盘信息')
    $plan.Add(('  型号：{0}' -f $r1.Model))
    $plan.Add(('  柱面：{0} - {1}；每柱面扇区：{2}（{3} 磁头 × {4} 扇区/磁道）' -f $r1.RangeStart, $r1.RangeEnd, $mapping.SectorsPerCylinder, $config.HeadsPerCylinder, $config.SectorsPerTrack))
    $plan.Add(('  总扇区：{0}（{1:N2} GiB）；逻辑扇区：{2} 字节' -f $mapping.CapacitySectors, (($mapping.CapacitySectors * [double]$config.LogicalSectorBytes) / 1073741824.0), $config.LogicalSectorBytes))
    $plan.Add(('  记录：{0} 条；明确错误：{1} 条；延迟中位数 {2} ms，P99 {3} ms' -f $r1.Count, $r1.ErrorCount, (Format-Num $r1.Stats.Median), (Format-Num $r1.Stats.P99)))
    $plan.Add('')
    $plan.Add('二、分区方案（1 MiB 对齐；GPT 首尾各保留 34 扇区）')
    if (-not $mapping.Verified) {
        $plan.Add('  未通过 LBA 映射验证，无法给出分区方案。')
        $plan.Add('  请确认报告从柱面 0 开始且记录连续；或在 Config.json 中手工确认映射。')
    } elseif ($partitions.Count -eq 0) {
        $plan.Add('  没有剩余连续可用空间。')
    } else {
        $shown = @($partitions | Where-Object Displayed)
        $hidden = @($partitions | Where-Object { -not $_.Displayed })
        $lastCylinderValue = [int64]$r1.Cylinder[$r1.Count - 1]
        if ($shown.Count -eq 0) {
            $plan.Add(('  没有容量占比 ≥ {0}% 的可用区间；可调小 Config.json 的 MinPartitionPercent。' -f $config.MinPartitionPercent))
        } else {
            $plan.Add('  序号 | 容量 | 占比 | 起始柱面 | 磁头 | 扇区 | 起始扇区号 | 总扇区数 | 结束柱面 | 结束磁头 | 结束扇区')
            for ($pi = 0; $pi -lt $shown.Count; $pi++) {
                $p = $shown[$pi]
                $plan.Add(('  {0,4} | {1,8:N2} GiB | {2,6:N2}% | {3,8} | {4,4} | {5,4} | {6,11} | {7,11} | {8,8} | {9,4} | {10,4}' -f $p.Partition, $p.SizeGiB, $p.Percent, $p.StartCylinder, $p.StartHead, $p.StartSector, $p.StartLba, $p.TotalSectors, $p.EndCylinder, $p.EndHead, $p.EndSector))
                $nextStart = if ($pi -lt ($shown.Count - 1)) { [int64]$shown[$pi + 1].StartCylinder } else { [int64]($lastCylinderValue + 1) }
                $between = @($iso.Ranges | Where-Object { $_.StartCylinder -gt $p.EndCylinder -and $_.StartCylinder -lt $nextStart } | Sort-Object StartCylinder)
                foreach ($rg in $between) {
                    $items = @($isoDetails | Where-Object { $_.RangeId -eq $rg.RangeId } | Sort-Object Cylinder)
                    if ($items.Count -le 5) {
                        $txt = (($items | ForEach-Object { '{0}={1}/{2}' -f $_.Cylinder, (Format-Num $_.Scan1Ms), (Format-Num $_.Scan2Ms) }) -join '，')
                    } else {
                        $mx1 = ($items | Measure-Object -Property Scan1Ms -Maximum).Maximum
                        $mx2 = ($items | Measure-Object -Property Scan2Ms -Maximum).Maximum
                        $txt = ('共 {0} 个异常点，最大 {1}/{2}' -f $items.Count, (Format-Num $mx1), (Format-Num $mx2))
                    }
                    $plan.Add(('       ↕ 隔离区#{0}（柱面 {1}-{2}，风险 {3}）：{4}' -f $rg.RangeId, $rg.StartCylinder, $rg.EndCylinder, $rg.Risk, $txt))
                }
            }
            $plan.Add('')
            $plan.Add(('  共显示 {0} 个分区，合计 {1:N2} GiB，占硬盘 {2:N2}%。' -f $shown.Count, (($shown | Measure-Object -Property SizeGiB -Sum).Sum), (($shown | Measure-Object -Property Percent -Sum).Sum)))
        }
        $plan.Add(('  显示阈值：分区容量 ≥ 硬盘总容量的 {0}%（Config.json: MinPartitionPercent）。' -f $config.MinPartitionPercent))
        if ($hidden.Count -gt 0) {
            $plan.Add(('  另有 {0} 个可用区间小于 {1}% 未显示（明细见 Partition_Scheme.csv 的 Displayed 列）。' -f $hidden.Count, $config.MinPartitionPercent))
        }
        $plan.Add('  说明：柱面/磁头/扇区为 255/63 传统几何；起始扇区号为绝对 LBA。')
        $plan.Add('        各分区彼此独立，不能跨越隔离区间合并。')
        $plan.Add(('  风险策略：{0}（strict=A/B/C/D 全部隔离；standard=A/B/D；errors-only=仅 A）。' -f $config.RiskProfile))
    }
    $plan.Add('')
    $plan.Add('三、必须避开的隔离区间')
    if ($iso.Ranges.Count -eq 0) { $plan.Add('  无。') }
    else {
        $plan.Add('  序号 | 柱面范围 | 风险 | 扫描1最大 | 扫描2最大 | 异常点 | 错误 | 区域P95 | 每侧扩展 | 成因')
        foreach ($rg in $iso.Ranges) {
            $plan.Add(('  {0,4} | {1,6}-{2,-6} | {3,4} | {4,8} | {5,8} | {6,5} | {7,4} | {8,7} | {9,4} MiB | {10}' -f $rg.RangeId, $rg.StartCylinder, $rg.EndCylinder, $rg.Risk, (Format-Num $rg.MaxMs1), (Format-Num $rg.MaxMs2), $rg.Points, $rg.ErrorPoints, (Format-Num $rg.RegionP95), $rg.GuardMiBEachSide, $rg.Reason))
        }
        $plan.Add('  说明：分区之间的空隙就是上表这些隔离区间；只有 3 处来自明确读取错误（A），')
        $plan.Add('        其余来自统计慢区（C）或单次/两次极端延迟（B/D），不是已确认的坏道。')
        $plan.Add('        “扫描1/2最大”为非错误记录的最大延迟；错误柱面的延迟见第五节。')
    }
    $plan.Add('')
    $plan.Add('四、隔离区异常柱面明细（柱面号 / 扫描1 ms / 扫描2 ms）')
    if ($isoDetails.Count -eq 0) { $plan.Add('  无。') }
    else {
        $plan.Add('  隔离区 | 柱面 | 扫描1 | 扫描2 | 标记')
        foreach ($d in $isoDetails) {
            $plan.Add(('  {0,6} | {1,6} | {2,6} | {3,6} | {4}' -f $d.RangeId, $d.Cylinder, (Format-Num $d.Scan1Ms), (Format-Num $d.Scan2Ms), $d.Class))
        }
    }
    $plan.Add('')
    $plan.Add('五、明确读取错误')
    if ($errList.Count -eq 0) { $plan.Add('  无。') }
    else {
        foreach ($e in $errList) { $plan.Add(('  柱面 {0}：{1} ms，{2}' -f $e.Cylinder, (Format-Num $e.Ms), $e.Text)) }
    }
    $plan.Add('')
    $plan.Add('六、校验与说明')
    $plan.Add(('  LBA 映射：{0}' -f $(if ($mapping.Verified) { '已验证（柱面 0 起、记录连续、SPC 一致、总扇区自洽）' } else { '未验证' })))
    if ($r1.Validation.Count -gt 0) {
        foreach ($item in $r1.Validation) { $plan.Add(('  [{0}] {1}' -f $item.Level, $item.Message)) }
    } else { $plan.Add('  结构校验：未发现问题。') }
    if ($null -ne $r2) {
        foreach ($item in $r2.Validation) { $plan.Add(('  [扫描2/{0}] {1}' -f $item.Level, $item.Message)) }
        foreach ($n in $compareNotes) { $plan.Add(('  [双报告] {0}' -f $n)) }
    }
    $plan.Add('  1. 统计异常不等于物理故障；隔离区间是保守建议，分区前请人工复核。')
    $plan.Add('  2. 本工具只读，不会修改磁盘或分区表；分区请用 DiskGenius/diskpart 人工执行。')
    $plan.Add('  3. 分区前请先备份重要数据。')
    $plan.Add(('  4. 完整耗时：{0:N1} 秒。' -f $script:RunClock.Elapsed.TotalSeconds))
    [System.IO.File]::WriteAllLines((Join-Path $out 'Partition_Plan.txt'), $plan.ToArray(), [System.Text.Encoding]::UTF8)

    if ($partitions.Count -gt 0) {
        Write-CsvFile -Path (Join-Path $out 'Partition_Scheme.csv') -Columns @('Partition', 'Recommended', 'StartCylinder', 'StartHead', 'StartSector', 'StartLba', 'TotalSectors', 'EndCylinder', 'EndHead', 'EndSector', 'EndLba', 'SizeGiB', 'Percent', 'Displayed') -Rows $partitions
    } else {
        [System.IO.File]::WriteAllText((Join-Path $out 'Partition_Scheme.csv'), 'Partition,Recommended,StartCylinder,StartHead,StartSector,StartLba,TotalSectors,EndCylinder,EndHead,EndSector,EndLba,SizeGiB,Percent', (New-Object System.Text.UTF8Encoding($true)))
    }

# ---- HTML（简洁版） ----
    $h = New-Object System.Text.StringBuilder
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode([string]$t) }
    [void]$h.Append('<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>DiskGuard 分区方案</title>')
    [void]$h.Append('<style>body{font-family:Segoe UI,Microsoft YaHei,sans-serif;margin:24px;color:#20242a}h1{font-size:20px;margin:0 0 4px}h2{font-size:15px;margin:22px 0 8px}table{border-collapse:collapse}th,td{border:1px solid #c8ccd2;padding:5px 9px;font-size:13px;text-align:left}th{background:#f0f2f5}.muted{color:#666;font-size:13px}.frag{color:#888}.iso td{background:#fff8e6;color:#7a4b00;font-size:12px;padding-left:24px}code{background:#f4f5f7;padding:1px 4px}</style></head><body>')
    [void]$h.Append('<h1>DiskGuard 分区方案</h1>')
    [void]$h.Append(('<p class="muted">生成时间 {0} · {1} · 只读分析，不会修改磁盘</p>' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $(if ($r2) { '双报告' } else { '单报告' })))
    [void]$h.Append('<h2>磁盘</h2><table>')
    [void]$h.Append(('<tr><th>型号</th><td>{0}</td></tr>' -f (& $enc $r1.Model)))
    [void]$h.Append(('<tr><th>柱面</th><td>{0} - {1}</td></tr>' -f $r1.RangeStart, $r1.RangeEnd))
    [void]$h.Append(('<tr><th>每柱面扇区</th><td>{0}（{1} 磁头 × {2} 扇区/磁道）</td></tr>' -f $mapping.SectorsPerCylinder, $config.HeadsPerCylinder, $config.SectorsPerTrack))
    [void]$h.Append(('<tr><th>总扇区</th><td>{0}（{1:N2} GiB）</td></tr>' -f $mapping.CapacitySectors, (($mapping.CapacitySectors * [double]$config.LogicalSectorBytes) / 1073741824.0)))
    [void]$h.Append(('<tr><th>记录 / 错误</th><td>{0} / {1}</td></tr>' -f $r1.Count, $r1.ErrorCount))
    [void]$h.Append('</table>')

    [void]$h.Append('<h2>分区方案（1 MiB 对齐；GPT 首尾各保留 34 扇区）</h2>')
    if (-not $mapping.Verified) {
        [void]$h.Append('<p class="muted">未通过 LBA 映射验证，无法给出分区方案。</p>')
    } elseif ($partitions.Count -eq 0) {
        [void]$h.Append('<p class="muted">没有剩余连续可用空间。</p>')
    } else {
        $shownHtml = @($partitions | Where-Object Displayed)
        $hiddenHtml = @($partitions | Where-Object { -not $_.Displayed })
        if ($shownHtml.Count -eq 0) {
            [void]$h.Append(('<p class="muted">没有容量占比 ≥ {0}% 的可用区间；可调小 Config.json 的 MinPartitionPercent。</p>' -f $config.MinPartitionPercent))
        } else {
            [void]$h.Append('<table><thead><tr><th>序号</th><th>容量</th><th>占比</th><th>起始柱面</th><th>磁头</th><th>扇区</th><th>起始扇区号</th><th>总扇区数</th><th>结束柱面</th><th>结束磁头</th><th>结束扇区</th></tr></thead><tbody>')
            $lastCylHtml = [int64]$r1.Cylinder[$r1.Count - 1]
            for ($pi = 0; $pi -lt $shownHtml.Count; $pi++) {
                $p = $shownHtml[$pi]
                [void]$h.Append(('<tr><td>{0}</td><td>{1:N2} GiB</td><td>{2:N2}%</td><td>{3}</td><td>{4}</td><td>{5}</td><td><code>{6}</code></td><td>{7}</td><td>{8}</td><td>{9}</td><td>{10}</td></tr>' -f $p.Partition, $p.SizeGiB, $p.Percent, $p.StartCylinder, $p.StartHead, $p.StartSector, $p.StartLba, $p.TotalSectors, $p.EndCylinder, $p.EndHead, $p.EndSector))
                $nextStart = if ($pi -lt ($shownHtml.Count - 1)) { [int64]$shownHtml[$pi + 1].StartCylinder } else { $lastCylHtml + 1 }
                $between = @($iso.Ranges | Where-Object { $_.StartCylinder -gt $p.EndCylinder -and $_.StartCylinder -lt $nextStart } | Sort-Object StartCylinder)
                foreach ($rg in $between) {
                    $items = @($isoDetails | Where-Object { $_.RangeId -eq $rg.RangeId } | Sort-Object Cylinder)
                    if ($items.Count -le 5) {
                        $txt = (($items | ForEach-Object { '{0}={1}/{2}' -f $_.Cylinder, (Format-Num $_.Scan1Ms), (Format-Num $_.Scan2Ms) }) -join '，')
                    } else {
                        $mx1 = ($items | Measure-Object -Property Scan1Ms -Maximum).Maximum
                        $mx2 = ($items | Measure-Object -Property Scan2Ms -Maximum).Maximum
                        $txt = ('共 {0} 个异常点，最大 {1}/{2}' -f $items.Count, (Format-Num $mx1), (Format-Num $mx2))
                    }
                    [void]$h.Append(('<tr class="iso"><td colspan="11">↕ 隔离区#{0}（柱面 {1}-{2}，风险 {3}）：{4}</td></tr>' -f $rg.RangeId, $rg.StartCylinder, $rg.EndCylinder, $rg.Risk, (& $enc $txt)))
                }
            }
            [void]$h.Append('</tbody></table>')
            [void]$h.Append(('<p class="muted">共显示 {0} 个分区，合计 {1:N2} GiB，占硬盘 {2:N2}%。</p>' -f $shownHtml.Count, (($shownHtml | Measure-Object -Property SizeGiB -Sum).Sum), (($shownHtml | Measure-Object -Property Percent -Sum).Sum)))
        }
        if ($hiddenHtml.Count -gt 0) {
            [void]$h.Append(('<p class="muted">另有 {0} 个可用区间小于 {1}% 未显示（明细见 Partition_Scheme.csv 的 Displayed 列）。</p>' -f $hiddenHtml.Count, $config.MinPartitionPercent))
        }
        [void]$h.Append(('<p class="muted">显示阈值：分区容量 ≥ 硬盘总容量的 {0}%。柱面/磁头/扇区为 255/63 传统几何；起始扇区号为绝对 LBA。</p>' -f $config.MinPartitionPercent))
    }

    [void]$h.Append('<h2>必须避开的隔离区间</h2><table><thead><tr><th>序号</th><th>柱面范围</th><th>风险</th><th>扫描1最大</th><th>扫描2最大</th><th>异常点</th><th>区域P95</th><th>每侧扩展</th><th>原因</th></tr></thead><tbody>')
    foreach ($rg in $iso.Ranges) {
        [void]$h.Append(('<tr><td>{0}</td><td>{1}-{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8} MiB</td><td>{9}</td></tr>' -f $rg.RangeId, $rg.StartCylinder, $rg.EndCylinder, $rg.Risk, (Format-Num $rg.MaxMs1), (Format-Num $rg.MaxMs2), $rg.Points, (Format-Num $rg.RegionP95), $rg.GuardMiBEachSide, (& $enc $rg.Reason)))
    }
    if ($iso.Ranges.Count -eq 0) { [void]$h.Append('<tr><td colspan="9" class="muted">无</td></tr>') }
    [void]$h.Append('</tbody></table>')

    [void]$h.Append('<h2>隔离区异常柱面明细</h2><table><thead><tr><th>隔离区</th><th>柱面</th><th>扫描1 ms</th><th>扫描2 ms</th><th>标记</th></tr></thead><tbody>')
    foreach ($d in $isoDetails) {
        [void]$h.Append(('<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f $d.RangeId, $d.Cylinder, (Format-Num $d.Scan1Ms), (Format-Num $d.Scan2Ms), $d.Class))
    }
    if ($isoDetails.Count -eq 0) { [void]$h.Append('<tr><td colspan="5" class="muted">无</td></tr>') }
    [void]$h.Append('</tbody></table>')

    [void]$h.Append('<h2>明确读取错误</h2>')
    if ($errList.Count -eq 0) { [void]$h.Append('<p class="muted">无。</p>') }
    else {
        [void]$h.Append('<table><thead><tr><th>柱面</th><th>耗时</th><th>错误信息</th></tr></thead><tbody>')
        foreach ($e in $errList) { [void]$h.Append(('<tr><td>{0}</td><td>{1} ms</td><td>{2}</td></tr>' -f $e.Cylinder, (Format-Num $e.Ms), (& $enc $e.Text))) }
        [void]$h.Append('</tbody></table>')
    }

    [void]$h.Append('<h2>说明</h2><ul class="muted">')
    [void]$h.Append(('<li>LBA 映射：{0}</li>' -f $(if ($mapping.Verified) { '已验证' } else { '未验证' })))
    foreach ($item in $r1.Validation) { [void]$h.Append(('<li>[{0}] {1}</li>' -f $item.Level, (& $enc $item.Message))) }
    if ($null -ne $r2) { foreach ($n in $compareNotes) { [void]$h.Append(('<li>[双报告] {0}</li>' -f (& $enc $n))) } }
    [void]$h.Append('<li>统计异常不等于物理故障；隔离区间是保守建议，分区前请人工复核。</li>')
    [void]$h.Append('<li>本工具只读，不会修改磁盘或分区表；分区请用 DiskGenius/diskpart 人工执行。</li>')
    [void]$h.Append('</ul>')
    [void]$h.Append('</body></html>')
    [System.IO.File]::WriteAllText((Join-Path $out 'Analysis_Report.html'), $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    # ---- 日志 ----
    $script:RunClock.Stop()
    Add-Log ('输出目录：' + $out)
    Add-Log ('总耗时：{0:N2}s' -f $script:RunClock.Elapsed.TotalSeconds)
    foreach ($u in $r1.Unparsed) { Add-Log ('未解析行 {0}: {1}' -f $u.LineNo, $u.Text) }
    [System.IO.File]::WriteAllLines((Join-Path $out 'Analysis.log'), $script:LogLines.ToArray(), [System.Text.Encoding]::UTF8)

    if (Test-Path -LiteralPath $script:FailureLog) { Remove-Item -LiteralPath $script:FailureLog -Force -ErrorAction SilentlyContinue }

    Write-Info ('完成。总耗时 {0:N1}s；输出：{1}' -f $script:RunClock.Elapsed.TotalSeconds, $out)
    Write-Info ('异常簇 {0} 个；隔离区间 {1} 个；LBA 映射 {2}。' -f $clusters.Count, $iso.Ranges.Count, $(if ($mapping.Verified) { '已验证' } else { '未验证' }))
    if (-not $mapping.Verified) { Write-Warn 'LBA 映射未通过验证，Partition_Plan.txt 只给出柱面范围。' }
    if (-not $NoBrowser) {
        try { Start-Process (Join-Path $out 'Analysis_Report.html') } catch { Write-Warn '无法自动打开浏览器，请手动打开 Analysis_Report.html。' }
    }
    return $out
}

# ---------------------------------------------------------------------------
# 入口（点源加载时不执行主流程，便于测试）
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    $exitCode = 0
    try {
        Invoke-DiskGuard -Report1 $Report1 -Report2 $Report2 -OutputRoot $OutputRoot -NoBrowser:$NoBrowser -SingleScan:$SingleScan -NoAccelerator:$NoAccelerator
    } catch {
        $message = $_.Exception.Message
        Write-Fail $message
        if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
        $logPath = $script:FailureLog
        if ([string]::IsNullOrWhiteSpace($logPath)) {
            $fallback = $PSScriptRoot
            if ([string]::IsNullOrWhiteSpace($fallback)) { $fallback = (Get-Location).Path }
            $fallbackOut = Join-Path $fallback 'Output'
            if (-not (Test-Path -LiteralPath $fallbackOut)) { New-Item -ItemType Directory -Path $fallbackOut -Force | Out-Null }
            $logPath = Join-Path $fallbackOut ("DiskGuard_failure_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        }
        try {
            [System.IO.File]::WriteAllLines($logPath, @(
                    ('DiskGuard v{0} failed: {1}' -f $script:Version, (Get-Date -Format o)),
                    ('Message: ' + $message),
                    ('Exception: ' + ($_.Exception.ToString())),
                    ('ScriptStackTrace: ' + $_.ScriptStackTrace)
                ), [System.Text.Encoding]::UTF8)
            Write-Warn ("错误日志已保存：{0}" -f $logPath)
        } catch {
            Write-Warn ("无法写入错误日志：{0}" -f $_.Exception.Message)
        }
        $exitCode = 1
    }
    if ($exitCode -ne 0 -and -not $NoPause -and $Host.Name -eq 'ConsoleHost') {
        try { [void](Read-Host '程序遇到错误，按 Enter 退出') } catch { }
    }
    exit $exitCode
}