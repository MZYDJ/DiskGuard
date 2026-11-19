#!/usr/bin/env bash
# 由 DiskGuard 根据 LVM 方案自动生成。
# 用法：sudo bash Linux_Setup.sh /dev/sdX
# 警告：会清空目标盘的分区表；执行前请确认设备名并备份数据。
set -euo pipefail

DISK="${1:-}"
if [ -z "$DISK" ]; then
  echo "用法: sudo bash Linux_Setup.sh /dev/sdX" >&2
  exit 1
fi
if [ ! -b "$DISK" ]; then
  echo "不是块设备: $DISK" >&2
  exit 1
fi

# 所有依赖必须在第一次写盘之前检查
for cmd in sgdisk partprobe udevadm blockdev lsblk grep findmnt head sort mountpoint wipefs pvcreate pvs vgs vgcreate lvcreate mkfs.ext4 blkid mkdir mount date tr sed cut; do
  command -v "$cmd" >/dev/null 2>&1 || { echo "缺少命令: $cmd" >&2; exit 1; }
done

EXPECTED_BYTES=9870336000
EXPECTED_LOGSS=512
EXPECTED_MODEL=''
MOUNT_POINT="/vol00/DISK_LVM"
ACTUAL_BYTES=$(blockdev --getsize64 "$DISK")
ACTUAL_LOGSS=$(blockdev --getss "$DISK")
if [ "$ACTUAL_BYTES" -ne "$EXPECTED_BYTES" ]; then
  echo "容量不匹配：期望 $EXPECTED_BYTES 字节，实际 $ACTUAL_BYTES 字节" >&2
  exit 1
fi
if [ "$ACTUAL_LOGSS" -ne "$EXPECTED_LOGSS" ]; then
  echo "逻辑扇区大小不匹配：期望 $EXPECTED_LOGSS，实际 $ACTUAL_LOGSS" >&2
  exit 1
fi

MODEL=$(lsblk -dno MODEL "$DISK" 2>/dev/null | sed "s/[[:space:]]*$//" || true)
SERIAL=$(lsblk -dno SERIAL "$DISK" 2>/dev/null | tr -d "[:space:]" || true)
WWN=$(lsblk -dno WWN "$DISK" 2>/dev/null | tr -d "[:space:]" || true)
echo "目标盘: $DISK"
echo "型号: $MODEL"
echo "序列号: ${SERIAL:-未知}"
echo "WWN: ${WWN:-未知}"
echo "容量: $ACTUAL_BYTES 字节"
echo "逻辑扇区: $ACTUAL_LOGSS 字节"
if [ -n "$EXPECTED_MODEL" ]; then
  if ! printf "%s" "$MODEL" | tr "[:upper:]" "[:lower:]" | grep -qF "$EXPECTED_MODEL"; then
    echo "型号不匹配：期望包含 $EXPECTED_MODEL，实际 $MODEL" >&2
    exit 1
  fi
fi

disk_in_use() {
  local disk="$1" src dev d
  while IFS= read -r src; do
    [ -z "$src" ] && continue
    for d in $(lsblk -sno NAME "$src" 2>/dev/null || true); do
      [ "/dev/$d" = "$disk" ] && return 0
    done
  done < <(findmnt -rn -o SOURCE 2>/dev/null | sort -u)
  while read -r dev _; do
    [ "$dev" = "Filename" ] && continue
    [ -z "$dev" ] && continue
    for d in $(lsblk -sno NAME "$dev" 2>/dev/null || true); do
      [ "/dev/$d" = "$disk" ] && return 0
    done
  done < /proc/swaps
  return 1
}
if disk_in_use "$DISK"; then
  echo "目标盘正在被当前系统使用（挂载或 swap），拒绝操作。" >&2
  exit 1
fi
if lsblk -no MOUNTPOINT "$DISK" 2>/dev/null | grep -q "[^[:space:]]"; then
  echo "目标盘存在挂载点，拒绝操作。" >&2
  exit 1
fi

STAMP=$(date +%Y%m%d_%H%M%S)
VG="vg_disk_${STAMP}"
if vgs "$VG" >/dev/null 2>&1; then
  echo "VG $VG 已存在，请先 vgrename 或修改脚本中的 VG 名。" >&2
  exit 1
fi
if mountpoint -q "$MOUNT_POINT"; then
  echo "挂载点 $MOUNT_POINT 已被占用，拒绝操作。" >&2
  exit 1
fi

echo "将创建以下分区（隔离带保持未分配）："
echo "  1: LBA 2048-8577023（4.09 GiB）"
echo "  2: LBA 10717184-19277823（4.08 GiB）"
echo "VG 名: $VG"
read -r -p "输入设备路径确认（$DISK）: " confirm
if [ "$confirm" != "$DISK" ]; then
  echo "设备路径不匹配，已取消。" >&2
  exit 1
fi
read -r -p "输入 YES 继续（将清空分区表）: " answer
[ "$answer" = "YES" ] || { echo "已取消"; exit 1; }

sgdisk --zap-all "$DISK"
sgdisk -n 1:2048:8577023 -t 1:8e00 -n 2:10717184:19277823 -t 2:8e00 "$DISK"
partprobe "$DISK"
udevadm settle
sgdisk -p "$DISK"
lsblk -o NAME,SIZE,TYPE,FSTYPE,MOUNTPOINT "$DISK"

part() { case "$DISK" in *[0-9]) echo "${DISK}p$1" ;; *) echo "${DISK}$1" ;; esac; }
wipefs -n "$(part 1)" "$(part 2)" || true
echo "以上为检测到的旧签名；继续执行 wipefs -a。"
wipefs -a "$(part 1)" "$(part 2)"
pvcreate "$(part 1)" "$(part 2)"
vgcreate "$VG" "$(part 1)" "$(part 2)"
lvcreate -l 100%FREE -n data "$VG"
mkfs.ext4 -L data "/dev/$VG/data"
mkdir -p "$MOUNT_POINT"
if mountpoint -q "$MOUNT_POINT"; then
  echo "挂载点 $MOUNT_POINT 已被占用，拒绝操作。" >&2
  exit 1
fi
mount -o noatime,errors=remount-ro "/dev/$VG/data" "$MOUNT_POINT"
UUID=$(blkid -s UUID -o value "/dev/$VG/data")
echo "完成。UUID: $UUID"
echo "挂载点: $MOUNT_POINT"
echo "开机挂载：把下面一行加入 /etc/fstab："
echo "  UUID=$UUID $MOUNT_POINT ext4 defaults,noatime,errors=remount-ro,nofail 0 2"
echo "注意：隔离带保持未分配；LVM 只合并空间，不修复慢区/坏道。"
