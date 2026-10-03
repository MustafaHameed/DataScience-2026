#!/usr/bin/env bash
# ch01_first_vm.sh -- create, snapshot, break and restore a virtual machine.
set -euo pipefail

# --- 1. Does this processor support virtualization at all? --------------
# vmx = Intel VT-x, svm = AMD-V. Chapter 4 explains what these flags buy.
# A count of 0 means the feature is absent or disabled in firmware.
grep -c -E '\<(vmx|svm)\>' /proc/cpuinfo
lscpu | grep -i 'virtuali[sz]ation' || echo "no such line reported"

# --- 2. Install the hypervisor and confirm it is usable -----------------
sudo apt-get install -y qemu-kvm libvirt-daemon-system virtinst \
                        libguestfs-tools cloud-image-utils
# virt-host-validate prints PASS/WARN/FAIL for every prerequisite.
sudo virt-host-validate qemu | head -20

# --- 3. Fetch a cloud image and give it a login -------------------------
# A cloud image is a disk file with an operating system already installed:
# encapsulation (section 1.3) before you have even started.
BASE=https://cloud-images.ubuntu.com/jammy/current
wget -nc -O ubuntu.img "$BASE/jammy-server-cloudimg-amd64.img"
cat > user-data <<'YAML'
#cloud-config
password: vss2026
chpasswd: { expire: False }
ssh_pwauth: True
YAML
cloud-localds seed.iso user-data

# --- 4. Create the virtual machine --------------------------------------
# Read the arguments as a parts list for an imaginary computer.
sudo cp ubuntu.img /var/lib/libvirt/images/vss1.qcow2
sudo qemu-img resize /var/lib/libvirt/images/vss1.qcow2 10G
sudo virt-install --name vss1 --memory 2048 --vcpus 2 \
  --disk /var/lib/libvirt/images/vss1.qcow2,format=qcow2 \
  --disk seed.iso,device=cdrom --os-variant ubuntu22.04 \
  --import --graphics none --noautoconsole

# --- 5. Look at it from the host ----------------------------------------
# dominfo is the guest as the hypervisor sees it: an allocation.
sudo virsh list --all
sudo virsh dominfo vss1

# --- 6. Snapshot it, then break it on purpose ---------------------------
sudo virsh snapshot-create-as vss1 clean "before I broke it"
sudo virsh console vss1    # log in as ubuntu / vss2026, then:
                           #   sudo rm -rf /etc/systemd  ; sudo reboot
                           # Ctrl-] returns you to the host.

# --- 7. Undo four minutes of damage in four seconds ---------------------
# No physical server supports this operation. That is section 1.3's point.
sudo virsh snapshot-revert vss1 clean --running
sudo virsh snapshot-list vss1

# --- 8. Tear down -------------------------------------------------------
sudo virsh destroy vss1 || true
sudo virsh undefine vss1 --remove-all-storage --snapshots-metadata
