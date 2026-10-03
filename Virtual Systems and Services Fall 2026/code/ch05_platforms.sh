#!/usr/bin/env bash
# ch05_platforms.sh -- one workload, two platforms, and what differs.
set -euo pipefail

# --- 1. What is this host already running, if anything? ----------------
systemd-detect-virt || echo "bare metal"
lsmod | grep -E '^(kvm|kvm_intel|kvm_amd|vboxdrv)' || true

# --- 2. Define a guest on KVM, from the Chapter 1 image. ---------------
sudo virsh dumpxml vss1 > vss1.xml
grep -E '<(domain|emulator|controller|interface|model)' vss1.xml | head -20
# Note the machine type: QEMU presents a specific, named virtual chipset.

# --- 3. Ask the guest what hardware it thinks it has. ------------------
sudo virsh start vss1 ; sleep 25
IP=$(sudo virsh domifaddr vss1 | awk '/ipv4/{split($4,a,"/"); print a[1]}')
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$IP" "$1"; }
G "sudo dmidecode -s system-manufacturer"
G "sudo dmidecode -s system-product-name"
G "lspci | head; systemd-detect-virt"

# --- 4. Build the same guest under VirtualBox. ------------------------
VBoxManage createvm --name vss2 --ostype Ubuntu_64 --register
VBoxManage modifyvm vss2 --memory 2048 --cpus 2 --nic1 nat
VBoxManage clonehd ubuntu.img vss2.vdi --format VDI
VBoxManage storagectl vss2 --name SATA --add sata
VBoxManage storageattach vss2 --storagectl SATA --port 0 \
           --type hdd --medium vss2.vdi
VBoxManage startvm vss2 --type headless

# --- 5. Compare what the two platforms claim to be. -------------------
# Each hypervisor identifies itself: the manufacturer string, the PCI
# device list and the CPUID hypervisor leaf all differ.
echo "KVM guest reports:"      ; G "sudo dmidecode -s system-manufacturer"
echo "VirtualBox guest reports: innotek GmbH (check in the guest console)"

# --- 6. The feature that is not in the free product. ------------------
# Live migration needs two hosts and shared storage. virsh offers it;
# VirtualBox does not. Chapter 17 is where this becomes the whole point.
sudo virsh migrate --help | head -5

# --- 7. Tear down. ----------------------------------------------------
sudo virsh destroy vss1 || true
VBoxManage controlvm vss2 poweroff || true
VBoxManage unregistervm vss2 --delete || true
