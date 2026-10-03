#!/usr/bin/env bash
# ch04_exits.sh -- see VMX, EPT and the exit rate with your own eyes.
set -euo pipefail

# --- 1. Does this processor have the features this chapter describes? ---
grep -o -m1 -E '\<(vmx|svm)\>' /proc/cpuinfo    # VT-x or AMD-V at all
grep -o -m1 -E '\<ept\>|\<npt\>' /proc/cpuinfo  # two-dimensional paging
grep -o -m1 '\<vpid\>' /proc/cpuinfo            # tagged TLB entries
# VT-d / AMD-Vi shows up as IOMMU groups once enabled in firmware:
ls /sys/kernel/iommu_groups 2>/dev/null | wc -l

# --- 2. Is nested virtualization on? Every later lab wants it. ---------
cat /sys/module/kvm_intel/parameters/nested 2>/dev/null \
  || cat /sys/module/kvm_amd/parameters/nested 2>/dev/null
# If that prints N or 0, see Appendix A for how to turn it on.

# --- 3. Start the guest from Chapter 1 and find its thread. ------------
sudo virsh start vss1 || true
sleep 25
PID=$(pgrep -f 'qemu.*vss1' | head -1)
echo "guest qemu pid: $PID"

# --- 4. Count exits by reason, for ten seconds of an idle guest. -------
# Every line is a different entry in the VM-execution controls.
sudo perf stat -e 'kvm:kvm_exit' -p "$PID" -- sleep 10
sudo perf record -e 'kvm:kvm_exit' -p "$PID" -- sleep 10 2>/dev/null
sudo perf script \
  | awk '{for(i=1;i<=NF;i++) if($i ~ /reason/) print $(i+1)}' \
  | sort | uniq -c | sort -rn | head -12

# --- 5. Now give the guest work that forces exits, and compare. --------
# Network I/O through an emulated device is an exit per register access.
in_guest () { ssh -o StrictHostKeyChecking=no "ubuntu@$1" "$2"; }
IP=$(sudo virsh domifaddr vss1 | awk '/ipv4/{split($4,a,"/"); print a[1]}')
in_guest "$IP" "ping -f -c 2000 10.0.2.2 >/dev/null 2>&1" &
sudo perf stat -e 'kvm:kvm_exit' -p "$PID" -- sleep 10
wait

# --- 6. Watch EPT doing its job: no exit for a guest page-table edit. --
# Fork a few thousand processes; under shadow paging this storms the host.
in_guest "$IP" "for i in \$(seq 2000); do /bin/true; done"
sudo perf stat -e 'kvm:kvm_exit' -p "$PID" -- sleep 5

# --- 7. Tear down ------------------------------------------------------
sudo virsh destroy vss1 || true
