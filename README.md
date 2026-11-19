# DiskGuard

> 分析 DiskGenius 全盘坏道扫描报告，输出可解释的隔离区间和带 CHS/LBA 参数的分区方案。
> Windows PowerShell 5.1 单文件工具，**严格只读**：不打开物理磁盘，不修改分区表。

DiskGuard 读取 DiskGenius 导出的 TXT 扫描报告，区分「明确读取错误」和「统计慢区」，
按风险等级给出建议隔离区间，并把磁盘剩余空间整理成可直接照抄到 DiskGenius / diskpart
的分区参数（起始柱面、磁头、扇区、起始扇区号、总扇区数）。

## 特性

- **只读**：只解析 TXT 并写出分析文件，不写磁盘、不碰分区表。
- **快**：60,802 条/份的双报告实测约 **7–8 秒**（旧版 v1.0.4 为 1017 秒）。
- **可解释**：每条隔离区间都带柱面范围、风险等级、最大延迟、异常点数量、区域 P95。
- **不过度隔离**：单个柱面即使两次都偏慢，只要未达临界幅度且不构成密集区域，只列入观察。
- **LBA 安全**：10 项结构证据 + 物理盘独立证据；任一不通过只给柱面范围。
- **可配置**：阈值、窗口、聚类、隔离带、风险策略、显示阈值都在 `Config.json`。
- **自带测试**：38 项自动化测试，覆盖规范要求的 20 类场景 + 边界条件 + 真实报告验收。

## 快速开始

1. 把整个目录解压到本地可写目录（不要放在只读介质或网络盘）。
2. 双击 `DiskGuard_启动.cmd`，或右键「打开」。
3. 选择第一次扫描报告；第二次报告可选（取消 = 单报告分析）。
4. 结果写入 `Output\<型号>_<时间>\`，窗口会保留并显示退出码。

命令行：

```powershell
# 双报告
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\DiskGuard.ps1 `
    -Report1 "D:\scan1.txt" -Report2 "D:\scan2.txt"

# 单报告
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\DiskGuard.ps1 `
    -Report1 "D:\scan1.txt" -SingleScan

# 不自动打开浏览器 / 不等待错误暂停 / 禁用加速内核
-NoBrowser -NoPause -NoAccelerator
```

## 输出

| 文件 | 内容 |
|---|---|
| `Partition_Plan.txt` | **主输出**：磁盘信息 → 分区方案（含容量、占比、CHS、LBA、总扇区数）→ 隔离区间 → 异常柱面明细 → 读取错误 → 校验说明 |
| `Partition_Scheme.csv` | 分区方案表格版，`Displayed` 列标记是否在方案中显示 |
| `Analysis_Report.html` | 同上的简洁网页版（无图表，纯表格） |
| `Isolation_Ranges.csv` | 隔离区间汇总（柱面 + LBA + 最大延迟 + 原因） |
| `Isolation_Anomalies.csv` | 每个隔离区里的异常柱面号与两次扫描值 |
| `Anomaly_Clusters.csv` | 异常簇 / 连续慢区的区域指标与风险等级 |
| `Anomaly_Details.csv` | 每个异常柱面的类别、延迟、局部阈值 |
| `Scan_Comparison.csv` | 每次扫描的统计基线（中位数、MAD、P99、阈值等） |
| `Linux_Setup.sh` | 映射验证通过时生成：Linux GPT 分区 + LVM 命令（VG 名从型号 + 时间生成） |
| `Analysis.log` | 阶段耗时、校验结果、未解析行、警告 |

`Cylinder_Comparison.csv`（5 MB 逐柱面对照）默认不生成，需要时把
`Config.json` 的 `WriteCylinderComparison` 设为 `true`。

### 分区方案示例

```
  序号 | 容量 | 占比 | 起始柱面 | 磁头 | 扇区 | 起始扇区号 | 总扇区数 | 结束柱面 | 结束磁头 | 结束扇区
     1 |   125.16 GiB |  26.87% |    40809 |   15 |   39 |   655597568 |   262483968 |    57147 |  237 |   50
       ↕ 隔离区#19（柱面 57148-57280，风险 D）：57214=593/1204
```

- 「占比」= 本分区容量 ÷ 硬盘总容量。
- 分区之间会列出夹在中间的隔离区：异常柱面 ≤5 个时逐条写 `柱面=扫描1/扫描2`，
  超过 5 个只写 `共 N 个异常点，最大 扫描1/扫描2`。

## Linux 分区与 LVM

LBA 映射验证通过时，`Partition_Plan.txt` 的第七节会直接列出**可逐条复制执行**的命令；
同时输出目录会多一个 `Linux_Setup.sh`（UTF-8 无 BOM、LF 换行），可一次性执行：

```bash
sudo bash Linux_Setup.sh /dev/sdX
```

报告里的命令按分区逐条列出，方便只执行其中一部分。脚本只使用 `Displayed=true` 的分区（同时满足 `MinPartitionGiB` 和 `MinPartitionPercent`）。脚本会：

1. 校验目标盘是块设备、容量与报告总扇区数一致、没有挂载点、不是当前系统盘；
2. 用 `sgdisk` 只在可用区间建 GPT 分区（类型 `8e00`），隔离带保持未分配；
3. `wipefs` → `pvcreate` → `vgcreate` → `lvcreate` → `mkfs.ext4`；
4. VG 名为 `vg_<型号>_<yyyymmdd_hhmm>`：型号取自报告（如 `ST3500418AS` → `st3500418as`），
   时间取执行时刻。

示例：

```bash
VG="vg_st3500418as_$(date +%Y%m%d_%H%M)"

sgdisk -n 1:2781184:44161023 -t 1:8e00 /dev/sdX
...
pvcreate /dev/sdX1
...
vgcreate "$VG" /dev/sdX1 ... /dev/sdX15
lvcreate -l 100%FREE -n data "$VG"
mkfs.ext4 -L data "/dev/$VG/data"
```

注意：脚本是**破坏性**的，会清空目标盘分区表；执行前请确认设备名并备份。
LVM 只合并空间，不修复慢区/坏道；隔离带仍然空着。

## 配置

`Config.json`（默认值）：

| 参数 | 默认 | 说明 |
|---|---:|---|
| `LogicalSectorBytes` | 512 | 逻辑扇区大小 |
| `MinRobustSigmaMs` | 5.0 | MAD 过小时的最小稳健尺度 |
| `GlobalK` / `ExtremeK` | 4 / 10 | 全盘 / 极端阈值系数 |
| `ExtremeRatioMin` | 10 | 风险临界值相对中位数的倍数 |
| `StrongRatioMin` | 6 | 「显著慢」相对中位数的倍数 |
| `LocalWindowSmall` / `LocalWindowLarge` | 50 / 200 | 局部窗口半径（柱面） |
| `LocalK` | 4 | 局部阈值系数 |
| `LocalMinSigmaMs` | 1.0 | 局部最小尺度 |
| `LocalRatioMin` | 1.5 | 局部相对倍数下限 |
| `LocalMinDeltaMs` | 100 | 局部最小绝对偏差，避免把微小抖动判成异常 |
| `LocalMinSamples` | 8 | 局部窗口最少有效样本 |
| `TightClusterCylinders` | 10 | 紧密聚类距离 |
| `MaxClusterSpan` | 200 | 单簇最大跨度（防链式合并） |
| `MinDensePoints` | 3 | 判定密集区域的最少异常点 |
| `ContextCylinders` | 10 | 区域背景统计的额外柱面数 |
| `RegionSlowRatio` | 1.5 | 区域整体偏慢判据 |
| `MinBackgroundSpan` | 20 | 连续慢区最小跨度 |
| `MinBackgroundFill` | 0.5 | 背景慢区最小填充率（异常柱面 ÷ 跨度） |
| `RiskProfile` | `strict` | 风险策略：`strict` / `standard` / `errors-only` |
| `MinPartitionGiB` | 2.0 | 小于该容量的区间不创建（碎片） |
| `MinPartitionPercent` | 5.0 | 分区方案和 Linux 脚本只使用容量占比 ≥ 该值的分区 |
| `HeadsPerCylinder` / `SectorsPerTrack` | 255 / 63 | CHS 几何 |
| `GuardErrorMiB` | 1024 | 明确错误每侧隔离带 |
| `GuardExtremeMiB` | 512 | 极端异常每侧隔离带 |
| `GuardRepeatedMiB` | 256 | 高密度异常区每侧隔离带 |
| `GuardBackgroundMiB` | 512 | 连续慢区每侧隔离带 |
| `RequireManualLbaConfirmation` | `false` | 是否强制人工确认 LBA 映射 |
| `RequireIndependentLbaEvidence` | `false` | 是否要求物理盘独立证据一致才输出 LBA |
| `LbaMappingConfirmed` | `false` | 人工确认标记 |
| `WriteCylinderComparison` | `false` | 是否输出 5 MB 逐柱面对照 CSV |

## 风险策略

隔离区间越多，分区方案被切得越碎。`RiskProfile` 控制哪些等级真正进入隔离：

| 策略 | 隔离等级 | 说明 |
|---|---|---|
| `strict`（默认） | A/B/C/D | 最保守 |
| `standard` | A/B/D | 跳过 470–670 ms 的密集慢区 |
| `errors-only` | A | 只隔离明确读取错误 |

风险等级：A 明确读取错误 · B 两次扫描均极端 / 区域持续偏慢 · C 高密度异常区 ·
D 单次极端 · E/F 观察。

## LBA 映射验证

DiskGenius 报告按柱面记录，分区工具要的是绝对扇区号。换算公式：

```
LBA = 柱面 × 16065 + 磁头 × 63 + (扇区 - 1)
```

只有以下结构证据全部通过才输出 LBA 和分区方案：

1. 扫描从柱面 0 开始；
2. 柱面编号连续；
3. 每柱面扇区数可确定；
4. 所有非末柱面的扇区数都等于该值（逐条检查，不只查总和）；
5. 最后柱面扇区数 ≤ 每柱面扇区数（完整柱面也允许）；
6. `(N-1) × S + 最后柱面扇区数 = 报告总扇区数`；
7. 声明的扫描末端等于实际最后柱面；
8. 记录数等于声明范围（防止报告尾部缺失仍被判为有效）；
9. 逻辑扇区大小有效；
10. 无重复柱面。

在此之上还有一层**独立证据**：通过只读 WMI 查找型号匹配的物理盘，比对
`Win32_DiskDrive.Size` 与报告计算出的容量。

- 找到且一致：`结构验证通过 + 物理盘独立证据一致`；
- 找不到：`结构验证通过（未找到匹配物理盘，未做独立核对）`；
- 找到但不一致：直接拒绝 LBA。

结论会明确写出验证等级，例如：

```
LBA 映射：已验证（结构验证通过（未找到匹配物理盘，未做独立核对））
```

任何一项不通过时只输出柱面范围，避免给出「看起来能用、实际会分错位置」的分区表。
`RequireManualLbaConfirmation=true` 可再加一道人工签字；
`RequireIndependentLbaEvidence=true` 要求物理盘独立证据必须一致，否则不输出 LBA。

## 算法流水线

```
解析与结构校验 → 单次扫描统计基线 → 全盘/极端/局部异常
→ 连续慢区与区域整体偏慢 → 双报告逐柱面对照 → 一维聚类
→ 风险分级 → 安全隔离区间 → LBA 验证 → 分区方案
```

- 全盘阈值 = `max(Median + 4σ, P99)`；极端阈值 = `max(Median + 10σ, P99.9)`。
- 局部阈值 = `max(局部中位数 + Kσ, 局部分位数)`，另需满足相对倍数与最小绝对偏差。
- 局部窗口统计由 `Add-Type` 编译的 C# 内核完成（失败时自动回退纯 PowerShell，结果一致）。

## 测试

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\DiskGuard.Tests.ps1
```

覆盖规范第 13 节要求的 20 类场景，并额外包含：

- 分区方案的 CHS/LBA/总扇区数一致性；
- 风险策略 `strict` / `errors-only` 的隔离数量差异；
- 加速内核与纯 PowerShell 回退的一致性；
- LBA 边界：最后柱面完整 / 不完整、非末柱面扇区数异常、声明末端不一致、尾部缺失、非零起始柱面；
- 区域慢区：稀疏异常点不合并、仅第二次扫描出现的慢区；
- 双报告记录数不一致时禁止正式 LBA；
- SPC 未知时不除零；
- Linux 脚本：验证通过时生成、未验证时不生成、BOM/关键命令检查；\n- 真实 ST3500418AS 双报告验收（记录数、3 条错误、8 个参考慢点、总扇区数、耗时、隔离区数量）。

全部通过时退出码为 0。加速内核不可用时，测试 22 会打印 SKIP，不判失败。

## 性能基准

本机实测（Windows 10.0.19041，PowerShell 5.1.19041.6093，8 核）：

| 场景 | 耗时 |
|---|---:|
| 双报告 60,802 × 2，加速内核 | 6.9–8.6 s |
| 单报告 60,802 条 | 约 5–6 s |
| 解析（每份） | 1.2–1.6 s |
| 异常检测 | 约 2.5 s |
| 聚类 | 约 0.7 s |
| 隔离区间 + 输出 | 约 2 s |

纯 PowerShell 回退在 60k 规模尚未完整计时（只在 400 条数据集上与加速路径做过一致性对比）；
Windows PE 环境也未实测。`Tests\Benchmark.ps1` 可复现上述数字，`-IncludeFallback` 会额外跑回退路径。

## 目录结构

```
DiskGuard.ps1                     主程序（Windows PowerShell 5.1）
DiskGuard_启动.cmd                双击启动入口
Config.json                       可调整参数
README.md                         本文件
Tests\DiskGuard.Tests.ps1         自动化测试（38 项）
Tests\Benchmark.ps1               性能基准（可选）
DiskGuard_算法设计与实现规范_v1.0.md   算法设计规范
```

> 仓库不包含真实硬盘扫描报告（每份约 10 MB，且属于用户数据）。
> 运行前请自备 DiskGenius 导出的 TXT 报告。

## 已知限制

1. 统计异常不等于物理故障；隔离区间是保守建议，必须人工审核后再分区。
2. 只能看到主机可见 LBA，无法读取硬盘固件已重映射的原始物理扇区。
3. 报告若缺少柱面范围或柱面编号不连续，LBA 映射会拒绝生成。
4. 单报告模式没有「重复性」证据，B 级（两次均极端）不会出现。
5. 纯 PowerShell 回退路径结果与加速路径一致，但 60k 记录规模会明显变慢。
6. 本工具不替代备份，也不替代硬盘可靠性评估。

## 旧版 v1.0.4 的问题

同一份报告实测：**1017 秒、1356 条候选、369 个隔离区**，且每柱面扇区数为空、LBA 未生成。根因：

1. 局部检测逐柱面重建 `List` 并排序，且没有分位数下限与最小绝对偏差；
2. 「两次都异常 → 直接隔离」没有幅度/密度约束；
3. 每柱面扇区数只接受「全盘完全一致」，最后一个不完整柱面导致 SPC 推断失败；
4. 逐行管道与重复属性访问。

v2.0.0 重写了整条流水线，详见 `DiskGuard_算法设计与实现规范_v1.0.md`。

## 许可证

[MIT](LICENSE)