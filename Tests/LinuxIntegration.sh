#!/usr/bin/env bash
# DiskGuard Linux 脚本端到端测试：loop device 上执行 GPT -> PV -> VG -> LV -> ext4。
# 用法（需要 root）：sudo bash Tests/LinuxIntegration.sh Tests/fixtures/Linux_Setup.sh
set -euo pipefail

FIXTURE="${1:-Tests/fixtures/Linux_Setup.sh}"
SIZE=9870336000
IMG=$(mktemp /tmp/diskguard.XXXXXX.img)
LOOP=""
VG=""
MOUNT=/mnt/data

cleanup() {
    set +e
    umount "$MOUNT" 2>/dev/null
    if [ -n "$VG" ]; then vgremove -f "$VG" >/dev/null 2>&1; fi
    if [ -n "$LOOP" ]; then losetup -d "$LOOP" 2>/dev/null; fi
    rm -f "$IMG"
}
trap cleanup EXIT

truncate -s "$SIZE" "$IMG"
# -P 让内核扫描 loop 设备上的分区
LOOP=$(losetup -P --find --show "$IMG")
echo "loop=$LOOP size=$SIZE"

printf '%s\nYES\n' "$LOOP" | bash "$FIXTURE" "$LOOP"

VG=$(vgs --noheadings -o vg_name | tr -d ' ' | tail -n1)
[ -n "$VG" ] || { echo "没有创建 VG" >&2; exit 1; }
LV=$(lvs --noheadings -o lv_name "$VG" | tr -d ' ' | tail -n1)
[ -n "$LV" ] || { echo "没有创建 LV" >&2; exit 1; }
DEV="/dev/$VG/$LV"
blkid "$DEV" | grep -q 'TYPE="ext4"' || { echo "文件系统不是 ext4" >&2; exit 1; }

umount "$MOUNT" 2>/dev/null || true
mkdir -p "$MOUNT"
mount "$DEV" "$MOUNT"
echo diskguard-test > "$MOUNT/diskguard_test.txt"
grep -q diskguard-test "$MOUNT/diskguard_test.txt"
umount "$MOUNT"
echo "OK: GPT -> PV -> VG -> LV -> ext4 verified on $LOOP"