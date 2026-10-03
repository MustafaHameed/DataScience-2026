#!/usr/bin/env bash
# ch11_microvm.sh -- 25 s, 125 ms and 100 ms, measured on one machine.
set -euo pipefail
cd ~ && mkdir -p fc && cd fc

# --- 1. Firecracker needs KVM and nothing else. ------------------------
[ -w /dev/kvm ] || { echo "no /dev/kvm -- see Chapter 4 step 1"; exit 1; }
ARCH=$(uname -m)
curl -sL -o firecracker \
  "https://github.com/firecracker-microvm/firecracker/releases/latest/\
download/firecracker-${ARCH}.tgz" || true
# If that fails, the release page lists a direct tarball; extract the
# firecracker binary beside this script.
chmod +x firecracker 2>/dev/null || true

# --- 2. A kernel and a root filesystem, loaded directly. --------------
# No bootloader and no firmware: the VMM places the kernel in memory.
curl -sL -o vmlinux \
  "https://s3.amazonaws.com/spec.ccfc.min/img/quickstart_guide/\
${ARCH}/kernels/vmlinux.bin"
curl -sL -o rootfs.ext4 \
  "https://s3.amazonaws.com/spec.ccfc.min/img/quickstart_guide/\
${ARCH}/rootfs/bionic.rootfs.ext4"

# --- 3. Describe the whole machine. Note how little there is. ---------
cat > vm.json <<'JSON'
{
  "boot-source": {
    "kernel_image_path": "vmlinux",
    "boot_args": "console=ttyS0 reboot=k panic=1 pci=off"
  },
  "drives": [{
    "drive_id": "rootfs", "path_on_host": "rootfs.ext4",
    "is_root_device": true, "is_read_only": false
  }],
  "machine-config": { "vcpu_count": 1, "mem_size_mib": 128 }
}
JSON
# pci=off is the chapter: there is no PCI bus to enumerate.

# --- 4. Boot it, and time to the login prompt. ------------------------
echo "=== firecracker ==="
time timeout 20 ./firecracker --no-api --config-file vm.json \
  2>&1 | grep -m1 -E 'login:|Welcome' || true

# --- 5. The same guest under full QEMU, for comparison. ---------------
echo "=== full qemu guest ==="
time timeout 90 sudo virsh start vss1 --console 2>&1 \
  | grep -m1 'login:' || true
sudo virsh destroy vss1 || true

# --- 6. And a container, for the other end of the axis. ---------------
echo "=== container ==="
time docker run --rm alpine:3.20 /bin/true

# --- 7. Memory footprint of the VMM process itself. -------------------
./firecracker --no-api --config-file vm.json &
FCPID=$!
sleep 3
grep -E '^(VmRSS|VmSize)' "/proc/$FCPID/status"
kill $FCPID 2>/dev/null || true
# Compare with the qemu process for vss1: typically 40-80x larger.

# --- 8. Kata, if the host has it: one flag, a whole VM. ---------------
# Under Kubernetes this is runtimeClassName: kata in the Pod spec.
command -v kata-runtime >/dev/null && kata-runtime check || \
  echo "kata not installed -- see Appendix A"

# --- 9. Tear down. ----------------------------------------------------
cd ~ && rm -rf fc
