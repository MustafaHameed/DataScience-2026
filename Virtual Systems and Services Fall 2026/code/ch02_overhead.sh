#!/usr/bin/env bash
# ch02_overhead.sh -- measure the cost of the layer, three ways.
set -euo pipefail
WORK='echo "scale=4000; 4*a(1)" | bc -l'   # seconds of pure arithmetic

# The guest's address, however libvirt chose to hand it out today.
ip_of () {
  sudo virsh domifaddr "$1" | awk '/ipv4/{split($4,a,"/"); print a[1]}'
}
in_guest () { ssh -o StrictHostKeyChecking=no "ubuntu@$(ip_of vss1)" "$1"; }

# --- 1. The baseline: no layer at all ----------------------------------
echo "=== native ==="
time bash -c "$WORK" > /dev/null

# --- 2. Type-1: KVM, hardware-assisted ---------------------------------
# vss1 is the guest built in Chapter 1's lab.
sudo virsh start vss1 || true
sleep 25
echo "=== type-1 (kvm) ==="
in_guest "time bash -c '$WORK' > /dev/null"

# --- 3. A container: no second kernel ----------------------------------
echo "=== container ==="
time docker run --rm ubuntu:22.04 bash -c \
     "apt-get -qq update && apt-get -qq install -y bc && $WORK" > /dev/null

# --- 4. Now the part that actually matters: I/O -------------------------
# Compare an emulated disk controller against a paravirtualized one.
# The guest must be stopped to change its disk bus.
sudo virsh destroy vss1
echo "=== emulated SATA controller ==="
sudo virt-xml vss1 --edit target=vda --disk bus=sata
sudo virsh start vss1 ; sleep 25
in_guest "sudo dd if=/dev/sda of=/dev/null bs=1M count=512 iflag=direct"

sudo virsh destroy vss1
echo "=== virtio (paravirtualized) ==="
sudo virt-xml vss1 --edit target=sda --disk bus=virtio
sudo virsh start vss1 ; sleep 25
in_guest "sudo dd if=/dev/vda of=/dev/null bs=1M count=512 iflag=direct"

# --- 5. Tear down -------------------------------------------------------
sudo virsh destroy vss1 || true
