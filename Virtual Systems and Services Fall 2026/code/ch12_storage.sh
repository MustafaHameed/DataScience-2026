#!/usr/bin/env bash
# ch12_storage.sh -- thin provisioning, snapshots and exhaustion.
set -euo pipefail

# --- 1. Virtual disk formats: what they cost on disk. ------------------
cd /tmp
qemu-img create -f raw   plain.raw  10G
qemu-img create -f qcow2 thin.qcow2 10G
ls -lh plain.raw thin.qcow2          # apparent size
du -h  plain.raw thin.qcow2          # actual size: this is the point
qemu-img info thin.qcow2 | head -6

# --- 2. A backing file: many disks, one shared base. ------------------
qemu-img create -f qcow2 -b thin.qcow2 -F qcow2 child1.qcow2
qemu-img create -f qcow2 -b thin.qcow2 -F qcow2 child2.qcow2
du -h child1.qcow2 child2.qcow2      # nearly nothing until written
qemu-img info child1.qcow2 | grep -i backing

# --- 3. Build a thin pool, deliberately over-subscribed. --------------
sudo truncate -s 2G /tmp/pv.img
LOOP=$(sudo losetup --find --show /tmp/pv.img)
sudo pvcreate "$LOOP" ; sudo vgcreate vsvg "$LOOP"
sudo lvcreate -L 1.5G --thinpool vspool vsvg
# Four 1 GB volumes out of a 1.5 GB pool: sigma = 2.7
for i in 1 2 3 4; do
  sudo lvcreate -V 1G --thin -n "vol$i" vsvg/vspool
done
sudo lvs vsvg -o lv_name,lv_size,data_percent

# --- 4. Put filesystems on them and write a little. -------------------
for i in 1 2 3 4; do
  sudo mkfs.ext4 -q "/dev/vsvg/vol$i"
  sudo mkdir -p "/mnt/v$i" && sudo mount "/dev/vsvg/vol$i" "/mnt/v$i"
  sudo dd if=/dev/urandom of="/mnt/v$i/data" bs=1M count=200 status=none
done
sudo lvs vsvg -o lv_name,data_percent    # pool filling up

# --- 5. Now walk off the cliff. ---------------------------------------
# Keep writing until the pool cannot satisfy a write.
sudo dd if=/dev/urandom of=/mnt/v1/more bs=1M count=900 || \
  echo "write failed -- the pool is full"
dmesg | tail -5 | grep -iE 'thin|read-only|I/O error' || true
mount | grep -E '/mnt/v[0-9]'            # look for "ro" on several at once

# --- 6. Snapshots and read amplification. -----------------------------
cd /tmp
qemu-img create -f qcow2 base.qcow2 1G
sudo modprobe nbd max_part=8 || true
for n in $(seq 1 8); do
  qemu-img create -f qcow2 -b "$( [ $n -eq 1 ] && echo base.qcow2 \
    || echo "snap$((n-1)).qcow2" )" -F qcow2 "snap$n.qcow2"
done
qemu-img info --backing-chain snap8.qcow2 | grep -c '^image:'
# Eight files, every one of which is required to read the newest.

# --- 7. Reclaim: discard is what returns space. -----------------------
sudo fstrim -v /mnt/v2 || echo "discard not supported on this stack"
sudo lvs vsvg -o lv_name,data_percent

# --- 8. Tear down. ----------------------------------------------------
for i in 1 2 3 4; do sudo umount "/mnt/v$i" || true; done
sudo lvremove -f vsvg || true ; sudo vgremove -f vsvg || true
sudo pvremove -f "$LOOP" || true ; sudo losetup -d "$LOOP" || true
rm -f /tmp/{plain.raw,thin.qcow2,child*.qcow2,base.qcow2,snap*.qcow2,pv.img}
