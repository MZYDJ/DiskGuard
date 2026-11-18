# DiskGuard v1.0.4 - DiskGenius bad-track report analyzer
# PowerShell 5.1; read-only. It never opens a physical disk for writing.
[CmdletBinding()]
param(
    [string]$Report1,
    [string]$Report2,
    [string]$OutputRoot,
    [switch]$NoBrowser,
    [switch]$NoPause,
    [switch]$InteractiveChild
)

$script:DGExplicitReport2 = $PSBoundParameters.ContainsKey('Report2')
$script:DGOutputRoot = $null
$script:DGFailureLog = $null

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

function Write-Info([string]$Message) { Write-Host "[DiskGuard] $Message" -ForegroundColor Cyan }
function Write-Warn([string]$Message) { Write-Host "[警告] $Message" -ForegroundColor Yellow }
function Write-Fail([string]$Message) { Write-Host "[错误] $Message" -ForegroundColor Red }

function Read-TextFile([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode.GetString($bytes,2,$bytes.Length-2) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes,2,$bytes.Length-2) }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8.GetString($bytes,3,$bytes.Length-3) }
    try { return [System.Text.Encoding]::UTF8.GetString($bytes) } catch { return [System.Text.Encoding]::Default.GetString($bytes) }
}

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

function Get-Percentile([double[]]$Values, [double]$P) {
    if ($Values.Count -eq 0) { return 0.0 }
    # Array.Sort avoids spawning a PowerShell pipeline for every local window.
    $a = [double[]]($Values.Clone())
    [Array]::Sort($a)
    if ($a.Count -eq 1) { return [double]$a[0] }
    $idx = ($a.Count - 1) * $P
    $lo = [int][Math]::Floor($idx); $hi = [int][Math]::Ceiling($idx)
    if ($lo -eq $hi) { return [double]$a[$lo] }
    return [double]$a[$lo] + (([double]$a[$hi] - [double]$a[$lo]) * ($idx - $lo))
}

function Get-Stats($Rows) {
    $v = @($Rows | Where-Object { -not $_.Error } | ForEach-Object { [double]$_.Ms })
    if ($v.Count -eq 0) { return [pscustomobject]@{Count=0;Median=0;MAD=0;Scale=1;P95=0;P99=0;P999=0;Max=0;Threshold=0;Severe=0} }
    $med = Get-Percentile $v 0.5
    $dev = @($v | ForEach-Object { [Math]::Abs($_ - $med) })
    $mad = Get-Percentile $dev 0.5
    $scale = [Math]::Max(1.0, 1.4826 * $mad)
    $p95 = Get-Percentile $v 0.95; $p99 = Get-Percentile $v 0.99; $p999 = Get-Percentile $v 0.999
    $max = ($v | Measure-Object -Maximum).Maximum
    $threshold = [Math]::Max(($med + 4.0*$scale), $p99)
    $severe = [Math]::Max(($med + 10.0*$scale), $p999)
    return [pscustomobject]@{Count=$v.Count;Median=$med;MAD=$mad;Scale=$scale;P95=$p95;P99=$p99;P999=$p999;Max=$max;Threshold=$threshold;Severe=$severe}
}

function Parse-Report([string]$Path, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "找不到报告文件：$Path" }
    $content = Read-TextFile $Path
    $lines = $content -split '(?:\r\n|\n|\r)'
    $model = $null
    if ($content -match '(?im)(?:被检测磁盘|检测磁盘|磁盘型号)\s*[:：]\s*(.+)$') { $model = $Matches[1].Trim() }
    $rangeStart = $null; $rangeEnd = $null; $rangeTotal = $null
    if ($content -match '(?im)检测范围\s*[:：]\s*柱面\s*(\d+)\s*[-～至]+\s*(\d+)') { $rangeStart=[int64]$Matches[1]; $rangeEnd=[int64]$Matches[2] }
    if ($content -match '(?im)完成时间\s*[:：]\s*(.+)$') { $scanTime=$Matches[1].Trim() } else { $scanTime=$null }
    $rows = New-Object System.Collections.Generic.List[object]
    $lineNo = 0
    foreach ($line in $lines) {
        $lineNo++
        if (($lineNo % 20000) -eq 0) { Write-Info ("报告文本解析进度：{0}/{1} 行；已识别 {2} 条记录。" -f $lineNo,$lines.Count,$rows.Count) }
        $s = $line.Trim()
        if ($s.Length -eq 0) { continue }
        # Expected columns: cylinder, head, sector, sector-count, elapsed-ms, repair-status, error.
        # Accept tab-separated and whitespace-aligned TXT exports.
        # DiskGenius TXT exports may use tabs, aligned spaces, or comma separators.
        # Accept all three; do not assume that columns are separated by a single tab.
        $m = [regex]::Match($s, '^\s*(\d+)[\t, ]+(\S+)[\t, ]+(\S+)[\t, ]+(\d+)[\t, ]+(\d+)(?:[\t, ]+(.*?))?\s*$')
        if (-not $m.Success) {
            # Some localized exports include explicit labels/units in the row.
            $m = [regex]::Match($s, '^\s*(\d+)\s+(.+?)\s+(.+?)\s+(\d+)\s+(\d+)\s*(?:ms|毫秒)?\s*(.*?)\s*$')
        }
        if (-not $m.Success) { continue }
        $cyl = [int64]$m.Groups[1].Value
        $head = $m.Groups[2].Value.Trim(); $sector = $m.Groups[3].Value.Trim()
        $count = [int64]$m.Groups[4].Value; $ms = [double]$m.Groups[5].Value
        $tail = $m.Groups[6].Value
        # Header rows and unrelated numeric lines are rejected by requiring a plausible sector count.
        if ($count -le 0 -or $count -gt 1000000 -or $ms -lt 0 -or $ms -gt 86400000) { continue }
        $isError = ($tail -match '(?i)unknown\s*error|error|错误|失败|坏道|未能读取|读取失败|无法读取|不可读取')
        # DiskGenius bad-track rows often have an explicit error in the last column.
        if ($tail -match '(?i)未处理\s*$' -and $s -match '(?i)Unknown Error') { $isError = $true }
        $rows.Add([pscustomobject]@{Cylinder=$cyl;Head=$head;Sector=$sector;SectorCount=$count;Ms=$ms;Tail=$tail;Error=[bool]$isError;Source=$Label;GlobalAnomaly=$false;Extreme=$false;LocalAnomaly=$false;LocalMedian=0.0;LocalThreshold=0.0})
    }
    if ($rows.Count -eq 0) {
        $sample = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 12) -join [Environment]::NewLine
        throw "未能从报告中解析出柱面数据。报告前12条非空行如下，便于定位格式差异：`n$sample`n文件：$Path"
    }
    $dups = @($rows | Group-Object Cylinder | Where-Object Count -gt 1)
    if ($dups.Count -gt 0) { throw "报告中存在重复柱面记录（例如柱面 $($dups[0].Name)），当前版本不支持这种记录结构。" }
    $sorted = @($rows | Sort-Object Cylinder)
    $counts = @($sorted | ForEach-Object { $_.SectorCount } | Select-Object -Unique)
    $spc = $null
    if ($counts.Count -eq 1) { $spc = [int64]$counts[0] }
    $stats = Get-Stats $sorted
    return [pscustomobject]@{Path=$Path;Label=$Label;Model=$model;ScanTime=$scanTime;RangeStart=$rangeStart;RangeEnd=$rangeEnd;Rows=$sorted;SectorsPerCylinder=$spc;Stats=$stats;LineCount=$lines.Count}
}

function Get-PercentileSorted([double[]]$SortedValues, [double]$P) {
    if ($SortedValues.Count -eq 0) { return 0.0 }
    if ($SortedValues.Count -eq 1) { return [double]$SortedValues[0] }
    $idx = ($SortedValues.Count - 1) * $P
    $lo = [int][Math]::Floor($idx); $hi = [int][Math]::Ceiling($idx)
    if ($lo -eq $hi) { return [double]$SortedValues[$lo] }
    return [double]$SortedValues[$lo] + (([double]$SortedValues[$hi] - [double]$SortedValues[$lo]) * ($idx - $lo))
}

function Add-AnomalyFlags($Report) {
    $stats = $Report.Stats
    $rows = @($Report.Rows)
    $total = $rows.Count

    # Global flags are O(n).
    foreach ($r in $rows) {
        $r.GlobalAnomaly = (-not $r.Error -and $r.Ms -gt $stats.Threshold)
        $r.Extreme = (-not $r.Error -and $r.Ms -ge $stats.Severe)
        $r.LocalAnomaly = $false
        $r.LocalMedian = 0.0
        $r.LocalThreshold = 0.0
    }

    # Local robust statistics. Avoid per-row Where-Object / ForEach-Object pipelines:
    # on 60k-cylinder reports those pipelines dominate runtime.
    for ($i=0; $i -lt $total; $i++) {
        if (($i % 5000) -eq 0 -and $i -gt 0) {
            Write-Info ("局部异常分析进度：{0}/{1}（{2:P0}）" -f $i,$total,($i/[double]$total))
        }
        if ($rows[$i].Error) { continue }

        $lo = [Math]::Max(0,$i-50)
        $hi = [Math]::Min($total-1,$i+50)
        $validCount = 0
        $windowValues = New-Object 'System.Collections.Generic.List[double]'
        for ($j=$lo; $j -le $hi; $j++) {
            if (-not $rows[$j].Error) {
                $windowValues.Add([double]$rows[$j].Ms)
                $validCount++
            }
        }
        if ($validCount -lt 15) { continue }

        [double[]]$sortedWindow = $windowValues.ToArray()
        [Array]::Sort($sortedWindow)
        $lm = Get-PercentileSorted $sortedWindow 0.5

        [double[]]$deviations = New-Object double[] $validCount
        for ($j=0; $j -lt $validCount; $j++) {
            $deviations[$j] = [Math]::Abs($sortedWindow[$j] - $lm)
        }
        [Array]::Sort($deviations)
        $lmad = Get-PercentileSorted $deviations 0.5

        $ls = [Math]::Max(($stats.Scale*0.5), [Math]::Max(1.0,1.4826*$lmad))
        $lt = $lm + 4.0*$ls
        $rows[$i].LocalMedian = $lm
        $rows[$i].LocalThreshold = $lt
        if ($rows[$i].Ms -gt $lt -and $rows[$i].Ms -gt ($lm*1.35)) {
            $rows[$i].LocalAnomaly = $true
        }
    }
    Write-Info ("局部异常分析完成：{0} 条记录。" -f $total)
    return $Report
}

function Build-CombinedCandidates($R1, $R2, [int]$ClusterDistance) {
    $map1=@{}; foreach($r in $R1.Rows){$map1[[string]$r.Cylinder]=$r}
    $map2=@{}; if($R2){foreach($r in $R2.Rows){$map2[[string]$r.Cylinder]=$r}}
    $cyls=@($map1.Keys | ForEach-Object {[int64]$_} | Sort-Object -Unique)
    $items=New-Object System.Collections.Generic.List[object]
    foreach($c in $cyls){
        $a=$map1[[string]$c]; $b=$null; if($R2 -and $map2.ContainsKey([string]$c)){$b=$map2[[string]$c]}
        $err1=[bool]$a.Error; $err2=($null -ne $b -and [bool]$b.Error)
        $anom1=([bool]$a.GlobalAnomaly -or [bool]$a.LocalAnomaly -or $err1)
        $anom2=($null -ne $b -and ([bool]$b.GlobalAnomaly -or [bool]$b.LocalAnomaly -or $err2))
        $ms1=[double]$a.Ms; $ms2=if($b){[double]$b.Ms}else{$null}
        $severity='正常'; $candidate=$false; $reason=''
        if($err1 -or $err2){$severity='明确读取错误';$candidate=$true;$reason='任意一次扫描报告读取错误'}
        elseif($b -and $anom1 -and $anom2){$severity='重复异常';$candidate=$true;$reason='两次扫描均出现异常'}
        elseif(($a.Extreme) -or ($b -and $b.Extreme)){$severity='单次极端异常';$candidate=$true;$reason='至少一次扫描达到极端延迟'}
        elseif($anom1 -or $anom2){$severity='单次异常';$candidate=$true;$reason='仅一次扫描出现全盘或局部异常'}
        if($candidate){$items.Add([pscustomobject]@{Cylinder=$c;Ms1=$ms1;Ms2=$ms2;Error1=$err1;Error2=$err2;Anomaly1=$anom1;Anomaly2=$anom2;Severity=$severity;Reason=$reason;Candidate=$true;NearCluster=$false})}
    }
    # Items are already ordered by cylinder. The nearest candidate must be one
    # of the two adjacent items, so avoid the previous O(n^2) full-list scan.
    for($i=0; $i -lt $items.Count; $i++){
        $item=$items[$i]
        if($item.Severity -eq '单次异常'){
            $nearest=[int64]::MaxValue
            if($i -gt 0){$nearest=[Math]::Min($nearest,[Math]::Abs([int64]$item.Cylinder-[int64]$items[$i-1].Cylinder))}
            if($i -lt ($items.Count-1)){$nearest=[Math]::Min($nearest,[Math]::Abs([int64]$items[$i+1].Cylinder-[int64]$item.Cylinder))}
            if($nearest -le $ClusterDistance){$item.NearCluster=$true;$item.Severity='邻近单次异常';$item.Reason+='；邻近其他异常点'}
            else {$item.Candidate=$false;$item.Severity='孤立单次异常（观察）';$item.Reason+='；远离其他异常，暂列观察'}
        }
    }
    return @($items | Sort-Object Cylinder)
}

function Build-Clusters($Candidates, [int]$Gap) {
    $active=@($Candidates | Where-Object Candidate | Sort-Object Cylinder)
    $clusters=New-Object System.Collections.Generic.List[object]
    if($active.Count -eq 0){return @()}
    $start=[int64]$active[0].Cylinder; $end=$start; $members=New-Object System.Collections.Generic.List[object]
    foreach($x in $active){
        if($members.Count -gt 0 -and ([int64]$x.Cylinder - $end) -gt ($Gap+1)){
            $clusters.Add([pscustomobject]@{Start=$start;End=$end;Count=$members.Count;Members=@($members.ToArray());MaxMs=(@($members | ForEach-Object { @($_.Ms1,$_.Ms2) | Where-Object {$_ -ne $null} } | Measure-Object -Maximum).Maximum)})
            $members=New-Object System.Collections.Generic.List[object];$start=[int64]$x.Cylinder
        }
        $members.Add($x);$end=[int64]$x.Cylinder
    }
    $clusters.Add([pscustomobject]@{Start=$start;End=$end;Count=$members.Count;Members=@($members.ToArray());MaxMs=(@($members | ForEach-Object { @($_.Ms1,$_.Ms2) | Where-Object {$_ -ne $null} } | Measure-Object -Maximum).Maximum)})
    return @($clusters.ToArray())
}

function Write-Csv($Data, [string]$Path) {
    # Export-Csv in Windows PowerShell 5.1 can throw a binder ArgumentException
    # when a property contains arrays of PSCustomObject (the cluster Members field).
    # Flatten every property to a scalar before handing it to Export-Csv.
    $flat = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in $Data) {
        $props = [ordered]@{}
        foreach ($property in $item.PSObject.Properties) {
            $value = $property.Value
            if ($null -eq $value) {
                $props[$property.Name] = ''
            }
            elseif ($property.Name -eq 'Members' -and $value -is [System.Array]) {
                $memberIds = New-Object 'System.Collections.Generic.List[string]'
                foreach ($member in $value) {
                    if ($null -ne $member -and $null -ne $member.PSObject.Properties['Cylinder']) {
                        $memberIds.Add([string]$member.Cylinder)
                    }
                    elseif ($null -ne $member) {
                        $memberIds.Add([string]$member)
                    }
                }
                $props[$property.Name] = ($memberIds -join ';')
            }
            elseif ($value -is [System.Array]) {
                $parts = New-Object 'System.Collections.Generic.List[string]'
                foreach ($part in $value) { if ($null -ne $part) { $parts.Add([string]$part) } }
                $props[$property.Name] = ($parts -join ';')
            }
            else {
                $props[$property.Name] = $value
            }
        }
        $flat.Add([pscustomobject]$props)
    }
    if ($flat.Count -eq 0) {
        [System.IO.File]::WriteAllText($Path,'',[System.Text.Encoding]::UTF8)
        return
    }
    $flat.ToArray() | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Main {
    $base = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($base)) { $base = (Get-Location).Path }
    $configPath = Join-Path $base 'Config.json'
    $config = [pscustomobject]@{LogicalSectorBytes=512;ClusterGapCylinders=10;ErrorGuardMiB=1024;ExtremeGuardMiB=512;RepeatedGuardMiB=256;SingleNearbyGuardMiB=256;LbaMappingConfirmed=$false;SectorsPerCylinderOverride=$null;DiskCapacitySectorsOverride=$null}
    if(Test-Path $configPath){$user=Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json;foreach($p in $user.PSObject.Properties){$config | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force}}
    if(-not $Report1){$Report1=Select-Report '选择第一次扫描报告（必选）' $true}
    if(-not $Report2 -and -not $script:DGExplicitReport2){$Report2=Select-Report '选择第二次扫描报告（可选，取消或留空则单报告分析）' $false}
    if(-not $OutputRoot){$OutputRoot=Join-Path $base 'Output'}
    $script:DGOutputRoot = $OutputRoot
    # Create a run log before parsing. Earlier builds only created Output after parsing,
    # so a format error left the operator with no diagnostic file at all.
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
    $stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:DGFailureLog = Join-Path $OutputRoot ("DiskGuard_failure_{0}.log" -f $stamp)
    [System.IO.File]::WriteAllLines($script:DGFailureLog, @('DiskGuard startup: ' + (Get-Date -Format o), 'Report1: ' + $Report1, 'Report2: ' + $Report2), [System.Text.Encoding]::UTF8)
    $runClock = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Info '正在读取并解析第一次扫描报告……'
    $stageClock = [System.Diagnostics.Stopwatch]::StartNew()
    $r1=Parse-Report $Report1 '扫描1'
    Write-Info ("第一次报告文本解析完成：{0} 条柱面记录；耗时 {1:N1} 秒。" -f $r1.Rows.Count,$stageClock.Elapsed.TotalSeconds)
    $stageClock.Restart()
    $r1=Add-AnomalyFlags $r1
    Write-Info ("第一次报告异常分析完成；耗时 {0:N1} 秒。" -f $stageClock.Elapsed.TotalSeconds)
    $r2=$null
    if($Report2){
        Write-Info '正在读取并解析第二次扫描报告……'
        $stageClock.Restart()
        $r2=Parse-Report $Report2 '扫描2'
        Write-Info ("第二次报告文本解析完成：{0} 条柱面记录；耗时 {1:N1} 秒。" -f $r2.Rows.Count,$stageClock.Elapsed.TotalSeconds)
        $stageClock.Restart()
        $r2=Add-AnomalyFlags $r2
        Write-Info ("第二次报告异常分析完成；耗时 {0:N1} 秒。" -f $stageClock.Elapsed.TotalSeconds)
    }
    $warnings=New-Object System.Collections.Generic.List[string]
    if($r2){
        if($r1.Model -and $r2.Model -and $r1.Model -ne $r2.Model){$warnings.Add('两份报告的磁盘标识不一致，禁止生成分区规划。')}
        if($r1.SectorsPerCylinder -and $r2.SectorsPerCylinder -and $r1.SectorsPerCylinder -ne $r2.SectorsPerCylinder){$warnings.Add('两份报告每柱面扇区数不一致，禁止生成分区规划。')}
        if($r1.Rows.Count -ne $r2.Rows.Count){$warnings.Add('两份报告记录数量不同，需人工核查扫描范围。')}
        if($r1.RangeStart -ne $r2.RangeStart -or $r1.RangeEnd -ne $r2.RangeEnd){$warnings.Add('两份报告扫描范围不同，禁止生成分区规划。')}
    }
    $spc=$r1.SectorsPerCylinder
    if($config.SectorsPerCylinderOverride){$spc=[int64]$config.SectorsPerCylinderOverride}
    $contiguous = $true
    for ($i=1; $i -lt $r1.Rows.Count; $i++) { if ([int64]$r1.Rows[$i].Cylinder -ne ([int64]$r1.Rows[$i-1].Cylinder + 1)) { $contiguous=$false; break } }
    $capacitySectors = $null
    if ($config.DiskCapacitySectorsOverride) { $capacitySectors = [int64]$config.DiskCapacitySectorsOverride }
    $contiguous2=$true
    if($r2){for($i=1;$i -lt $r2.Rows.Count;$i++){if([int64]$r2.Rows[$i].Cylinder -ne ([int64]$r2.Rows[$i-1].Cylinder+1)){$contiguous2=$false;break}}}
    $mappingOK=[bool]$config.LbaMappingConfirmed -and ($null -ne $spc) -and ($spc -gt 0) -and ($null -ne $capacitySectors) -and ($capacitySectors -gt 0) -and $contiguous -and $contiguous2 -and ($r1.Rows[0].Cylinder -eq 0)
    $stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeModel=if($r1.Model){($r1.Model -replace '[^\w\-]','_')}else{'UnknownDisk'}
    $out=Join-Path $OutputRoot ("{0}_{1}" -f $safeModel,$stamp)
    New-Item -ItemType Directory -Path $out -Force | Out-Null
    Write-Info '正在合并两次扫描的异常候选……'
    $stageClock.Restart()
    $candidates=Build-CombinedCandidates $r1 $r2 ([int]$config.ClusterGapCylinders)
    Write-Info ("异常候选合并完成：{0} 条；耗时 {1:N1} 秒。" -f $candidates.Count,$stageClock.Elapsed.TotalSeconds)
    Write-Info '正在进行异常区间聚类……'
    $stageClock.Restart()
    $clusters=Build-Clusters $candidates ([int]$config.ClusterGapCylinders)
    Write-Info ("异常区间聚类完成：{0} 个；耗时 {1:N1} 秒。" -f $clusters.Count,$stageClock.Elapsed.TotalSeconds)
    $guardMap=@{ '明确读取错误'=[double]$config.ErrorGuardMiB; '重复异常'=[double]$config.RepeatedGuardMiB; '单次极端异常'=[double]$config.ExtremeGuardMiB; '邻近单次异常'=[double]$config.SingleNearbyGuardMiB }
    $ranges=New-Object System.Collections.Generic.List[object]
    foreach($cl in $clusters){
        $sev='重复异常'; if(@($cl.Members | Where-Object Severity -eq '明确读取错误').Count -gt 0){$sev='明确读取错误'} elseif(@($cl.Members | Where-Object Severity -eq '单次极端异常').Count -gt 0){$sev='单次极端异常'} elseif(@($cl.Members | Where-Object Severity -eq '邻近单次异常').Count -gt 0){$sev='邻近单次异常'}
        $guard=$guardMap[$sev]; if(-not $guard){$guard=[double]$config.RepeatedGuardMiB}
        $ranges.Add([pscustomobject]@{StartCylinder=$cl.Start;EndCylinder=$cl.End;AnomalyCount=$cl.Count;Risk=$sev;GuardEachSideMiB=$guard;LbaStart=$null;LbaEnd=$null;Status='待确认LBA映射'})
    }
    if($mappingOK){
        $alignment=[int64][Math]::Ceiling(1MB/[double]$config.LogicalSectorBytes)
        $firstUsable=[int64]([Math]::Ceiling(34.0/$alignment)*$alignment)
        $lastUsable=[int64]([Math]::Floor(($capacitySectors-34)/[double]$alignment)*$alignment-1)
        foreach($rg in $ranges){
            $rawStart=[int64]$rg.StartCylinder*$spc; $rawEnd=([int64]$rg.EndCylinder+1)*$spc-1
            $guardSectors=[int64][Math]::Ceiling(($rg.GuardEachSideMiB*1MB)/[double]$config.LogicalSectorBytes)
            $candidateStart=[Math]::Max($firstUsable,($rawStart-$guardSectors))
            $candidateEnd=[Math]::Min($lastUsable,($rawEnd+$guardSectors))
            $rg.LbaStart=[int64]([Math]::Floor($candidateStart/[double]$alignment)*$alignment)
            $rg.LbaEnd=[int64]([Math]::Min($lastUsable,([Math]::Ceiling(($candidateEnd+1)/[double]$alignment)*$alignment)-1))
            $rg.Status='按1MiB向外对齐的候选隔离范围；需人工审核GPT边界'
        }
        # Merge overlapping expanded LBA ranges.
        $sorted=@($ranges | Sort-Object LbaStart);$merged=New-Object System.Collections.Generic.List[object]
        foreach($rg in $sorted){if($merged.Count -gt 0 -and $rg.LbaStart -le ($merged[$merged.Count-1].LbaEnd+1)){$last=$merged[$merged.Count-1];$last.LbaEnd=[Math]::Max($last.LbaEnd,$rg.LbaEnd);$last.StartCylinder=[Math]::Min($last.StartCylinder,$rg.StartCylinder);$last.EndCylinder=[Math]::Max($last.EndCylinder,$rg.EndCylinder);$last.Risk='合并隔离区';$last.AnomalyCount += $rg.AnomalyCount}else{$merged.Add($rg)}}
        $ranges=$merged
    }
    $scanSummary=@([pscustomobject]@{Scan='扫描1';Model=$r1.Model;Time=$r1.ScanTime;Records=$r1.Rows.Count;MedianMs=$r1.Stats.Median;MAD=$r1.Stats.MAD;P95=$r1.Stats.P95;P99=$r1.Stats.P99;P999=$r1.Stats.P999;ThresholdMs=$r1.Stats.Threshold;SevereMs=$r1.Stats.Severe})
    if($r2){$scanSummary+= [pscustomobject]@{Scan='扫描2';Model=$r2.Model;Time=$r2.ScanTime;Records=$r2.Rows.Count;MedianMs=$r2.Stats.Median;MAD=$r2.Stats.MAD;P95=$r2.Stats.P95;P99=$r2.Stats.P99;P999=$r2.Stats.P999;ThresholdMs=$r2.Stats.Threshold;SevereMs=$r2.Stats.Severe}}
    Write-Info '正在导出统计、异常和隔离区 CSV……'
    $stageClock.Restart()
    Write-Csv $scanSummary (Join-Path $out 'Scan_Comparison.csv')
    Write-Csv $candidates (Join-Path $out 'Anomaly_Details.csv')
    Write-Csv $clusters (Join-Path $out 'Anomaly_Clusters.csv')
    Write-Csv $ranges (Join-Path $out 'Isolation_Ranges.csv')
    Write-Info ("统计与异常 CSV 导出完成；耗时 {0:N1} 秒。" -f $stageClock.Elapsed.TotalSeconds)
    # Build the second-scan lookup once. The previous implementation ran
    # Where-Object over all 60k rows for every first-scan row (O(n^2)),
    # which could take tens of minutes or appear frozen.
    $allRows=New-Object System.Collections.Generic.List[object]
    $r2Lookup=$null
    if($r2){
        Write-Info '正在建立第二次扫描柱面索引……'
        $stageClock.Restart()
        $r2Lookup=@{}
        foreach($b in $r2.Rows){$r2Lookup[[string]$b.Cylinder]=$b}
        Write-Info ("第二次扫描索引完成：{0} 条；耗时 {1:N1} 秒。" -f $r2Lookup.Count,$stageClock.Elapsed.TotalSeconds)
    }
    Write-Info '正在生成逐柱面对比明细……'
    $stageClock.Restart()
    for($i=0;$i -lt $r1.Rows.Count;$i++){
        $r=$r1.Rows[$i]
        if($r2Lookup){
            $b=$r2Lookup[[string]$r.Cylinder]
            $allRows.Add([pscustomobject]@{Cylinder=$r.Cylinder;Scan1Ms=$r.Ms;Scan1Error=$r.Error;Scan1Global=$r.GlobalAnomaly;Scan1Local=$r.LocalAnomaly;Scan2Ms=$(if($b){$b.Ms}else{$null});Scan2Error=$(if($b){$b.Error}else{$null});Scan2Global=$(if($b){$b.GlobalAnomaly}else{$null});Scan2Local=$(if($b){$b.LocalAnomaly}else{$null})})
        } else {
            $allRows.Add([pscustomobject]@{Cylinder=$r.Cylinder;Scan1Ms=$r.Ms;Scan1Error=$r.Error;Scan1Global=$r.GlobalAnomaly;Scan1Local=$r.LocalAnomaly})
        }
        if(($i+1)%10000 -eq 0){Write-Info ("逐柱面对比进度：{0}/{1}" -f ($i+1),$r1.Rows.Count)}
    }
    Write-Info ("逐柱面对比完成；耗时 {0:N1} 秒。" -f $stageClock.Elapsed.TotalSeconds)
    Write-Info '正在导出逐柱面对比 CSV……'
    $stageClock.Restart()
    Write-Csv $allRows (Join-Path $out 'Cylinder_Comparison.csv')
    Write-Info ("逐柱面对比 CSV 导出完成；耗时 {0:N1} 秒。" -f $stageClock.Elapsed.TotalSeconds)
    $plan=New-Object System.Collections.Generic.List[string]
    $plan.Add('DiskGuard GPT 分区规划（仅供人工审核；本程序不会修改磁盘）')
    $plan.Add("磁盘标识：$($r1.Model)");$plan.Add("逻辑扇区大小配置：$($config.LogicalSectorBytes) bytes");$plan.Add("每柱面扇区数：$spc");$plan.Add("配置的磁盘总扇区数：$capacitySectors")
    if(-not $mappingOK){$plan.Add('状态：未生成可用 LBA 边界。请确认报告从柱面0开始且记录连续，并在 Config.json 中填写磁盘总扇区数、逻辑扇区大小、每柱面扇区数，核实映射后设置 LbaMappingConfirmed=true。')}
    elseif($warnings.Count -gt 0){$plan.Add('状态：报告校验存在问题，以下范围仅为候选，不得用于分区操作。')}
    else{$plan.Add('状态：已根据配置计算候选隔离范围；仍须人工核对磁盘容量、GPT头尾保留区和1MiB对齐。')}
    foreach($w in $warnings){$plan.Add("警告：$w")}
    foreach($rg in $ranges){$plan.Add(("柱面 {0}-{1} | LBA {2}-{3} | 风险 {4} | 两侧扩展 {5} MiB | {6}" -f $rg.StartCylinder,$rg.EndCylinder,$rg.LbaStart,$rg.LbaEnd,$rg.Risk,$rg.GuardEachSideMiB,$rg.Status))}
    if($mappingOK -and $warnings.Count -eq 0){
        $alignment=[int64][Math]::Ceiling(1MB/[double]$config.LogicalSectorBytes)
        $firstUsable=[int64]([Math]::Ceiling(34.0/$alignment)*$alignment)
        $lastUsable=[int64]([Math]::Floor(($capacitySectors-34)/[double]$alignment)*$alignment-1)
        $cursor=$firstUsable; $extentNo=0
        foreach($rg in @($ranges | Sort-Object LbaStart)){
            if($rg.LbaStart -gt $cursor){$extentNo++;$plan.Add(("可用分区候选{0}：LBA {1}-{2}" -f $extentNo,$cursor,($rg.LbaStart-1)))}
            $cursor=[Math]::Max($cursor,($rg.LbaEnd+1))
        }
        if($cursor -le $lastUsable){$extentNo++;$plan.Add(("可用分区候选{0}：LBA {1}-{2}" -f $extentNo,$cursor,$lastUsable))}
        if($extentNo -eq 0){$plan.Add('没有剩余的连续可用分区空间。')}
        $plan.Add('说明：每个可用候选区间是独立的连续空间；多个区间不能合并成一个跨越隔离区的分区。GPT首尾保留空间和1MiB对齐已按配置作保守估算。')
    }
    [System.IO.File]::WriteAllLines((Join-Path $out 'Partition_Plan.txt'),$plan,[System.Text.Encoding]::UTF8)
    $htmlRows=foreach($x in $scanSummary){"<tr><td>$($x.Scan)</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$x.Model))</td><td>$($x.Records)</td><td>$([Math]::Round($x.MedianMs,1))</td><td>$([Math]::Round($x.P95,1))</td><td>$([Math]::Round($x.P99,1))</td><td>$([Math]::Round($x.ThresholdMs,1))</td></tr>"}
    $candRows=foreach($x in $candidates){"<tr><td>$($x.Cylinder)</td><td>$($x.Ms1)</td><td>$($x.Ms2)</td><td>$($x.Severity)</td><td>$($x.Reason)</td></tr>"}
    $warnHtml=if($warnings.Count){($warnings | ForEach-Object { '<li>'+[System.Net.WebUtility]::HtmlEncode($_)+'</li>' }) -join ''}else{'<li>未发现报告身份或范围校验警告。</li>'}
    $html=@"
<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>DiskGuard 分析报告</title><style>body{font-family:Segoe UI,Microsoft YaHei,sans-serif;margin:28px;color:#20242a;background:#f6f7f9}main{max-width:1200px;margin:auto}.card{background:white;border:1px solid #ddd;border-radius:10px;padding:18px;margin:14px 0}h1{margin-bottom:4px}small,.muted{color:#68717d}table{border-collapse:collapse;width:100%;font-size:14px}th,td{padding:9px;border-bottom:1px solid #e5e7eb;text-align:left}th{background:#f0f2f5}.warn{color:#9a5700}.pill{display:inline-block;padding:4px 8px;border-radius:5px;background:#fff2d8}.note{line-height:1.65}</style></head><body><main><h1>DiskGuard 分析报告</h1><div class="muted">生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')　|　模式：$(if($r2){'双报告'}else{'单报告'})</div><section class="card"><h2>扫描统计</h2><table><thead><tr><th>扫描</th><th>磁盘标识</th><th>记录数</th><th>中位数 ms</th><th>P95 ms</th><th>P99 ms</th><th>异常阈值 ms</th></tr></thead><tbody>$($htmlRows -join "`n")</tbody></table></section><section class="card"><h2>校验与警告</h2><ul class="note">$warnHtml</ul><p class="note">LBA映射确认状态：<b>$(if($mappingOK){'配置已确认（仍需人工核验）'}else{'未确认，禁止把候选范围作为正式分区边界'})</b></p></section><section class="card"><h2>异常候选明细</h2><table><thead><tr><th>柱面</th><th>扫描1 ms</th><th>扫描2 ms</th><th>分类</th><th>原因</th></tr></thead><tbody>$($candRows -join "`n")</tbody></table></section><section class="card"><h2>隔离区摘要</h2><p>候选异常点：$($candidates.Count)；聚类区域：$($clusters.Count)；最终隔离区：$($ranges.Count)。</p><p class="note">本工具依据主机可见LBA扫描结果识别异常，不能读取已被硬盘固件重映射的原始物理扇区，也不能保证未来不会产生新坏道。分析结果不替代备份或硬盘可靠性评估。</p></section><footer class="muted">DiskGuard · 只读分析工具 · 不执行任何磁盘写入</footer></main></body></html>
"@
    [System.IO.File]::WriteAllText((Join-Path $out 'Analysis_Report.html'),$html,[System.Text.Encoding]::UTF8)
    $runClock.Stop()
    $log=@("DiskGuard run: $(Get-Date -Format o)","Report1: $Report1","Report2: $Report2","Output: $out","Scan1 rows: $($r1.Rows.Count)","Scan2 rows: $(if($r2){$r2.Rows.Count}else{0})","Candidates: $($candidates.Count)","Clusters: $($clusters.Count)","Isolation ranges: $($ranges.Count)","Total elapsed seconds: $([Math]::Round($runClock.Elapsed.TotalSeconds,1))","LbaMappingConfirmed: $mappingOK") + @($warnings)
    [System.IO.File]::WriteAllLines((Join-Path $out 'Analysis.log'),$log,[System.Text.Encoding]::UTF8)
    if(Test-Path -LiteralPath $script:DGFailureLog){Remove-Item -LiteralPath $script:DGFailureLog -Force -ErrorAction SilentlyContinue}
    Write-Info "分析完成。总耗时：$([Math]::Round($runClock.Elapsed.TotalSeconds,1)) 秒。输出目录：$out"
    Write-Info "异常候选 $($candidates.Count) 个；聚类 $($clusters.Count) 个；隔离区 $($ranges.Count) 个。"
    if(-not $mappingOK){Write-Warn 'LBA映射尚未由操作者确认，因此分区规划文件不会给出可直接使用的分区边界。'}
    if(-not $NoBrowser){try{Start-Process (Join-Path $out 'Analysis_Report.html')}catch{Write-Warn '无法自动打开浏览器，请手动打开 Analysis_Report.html。'}}
}

try {
    Main
} catch {
    $message = $_.Exception.Message
    Write-Fail $message
    if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray }
    $logPath = $script:DGFailureLog
    if ([string]::IsNullOrWhiteSpace($logPath)) {
        $fallback = $PSScriptRoot
        if ([string]::IsNullOrWhiteSpace($fallback)) { $fallback = (Get-Location).Path }
        $logPath = Join-Path $fallback ("DiskGuard_failure_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    try {
        $details = @(
            'DiskGuard failed: ' + (Get-Date -Format o),
            'Message: ' + $message,
            'Exception: ' + ($_.Exception.ToString()),
            'ScriptStackTrace: ' + $_.ScriptStackTrace
        )
        [System.IO.File]::WriteAllLines($logPath, $details, [System.Text.Encoding]::UTF8)
        Write-Warn ("错误日志已保存：{0}" -f $logPath)
    } catch {
        Write-Warn ("无法写入错误日志：{0}" -f $_.Exception.Message)
    }
    if (-not $NoPause -and $Host.Name -eq 'ConsoleHost') {
        try { [void](Read-Host '程序遇到错误，按 Enter 退出（错误日志已尽可能保存）') } catch {}
    }
    exit 1
}
