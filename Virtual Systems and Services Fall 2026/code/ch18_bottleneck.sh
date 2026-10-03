#!/usr/bin/env bash
# ch18_bottleneck.sh -- CPU contention, memory pressure and I/O queueing.
set -euo pipefail
sudo apt-get install -y sysstat fio stress-ng
ip_of () {
  sudo virsh domifaddr "$1" | awk '/ipv4/{split($4,a,"/"); print a[1]}'
}
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$(ip_of "$1")" "$2"; }

# --- 1. Baseline, with the host quiet. --------------------------------
sudo virsh start vss1 || true ; sleep 25
G vss1 "sudo apt-get -qq install -y sysstat fio stress-ng"
echo "=== baseline ==="
G vss1 "mpstat 2 3 | tail -2"           # %steal should be near zero
sudo virsh domstats vss1 --vcpu | grep -E 'wait|delay' || true

# --- 2. Bottleneck A: CPU contention from the neighbours. -------------
# Other guests take the cores. The complaining guest changes nothing.
for n in 2 3 4; do
  sudo virt-clone --original vss1 --name "n$n" \
    --file "/var/lib/libvirt/images/n$n.qcow2" --auto-clone >/dev/null
  sudo virsh setvcpus "n$n" "$(nproc)" --config --maximum
  sudo virsh setvcpus "n$n" "$(nproc)" --config
  sudo virsh start "n$n"
done
sleep 40
for n in 2 3 4; do
  G "n$n" "nohup stress-ng --cpu \$(nproc) --timeout 120s >/dev/null 2>&1 &"
done
sleep 10
echo "=== A: CPU contention ==="
G vss1 "mpstat 2 3 | tail -2"           # %steal climbs; %usr does NOT
sudo virsh domstats vss1 --vcpu | grep -E 'wait|delay' || true
# The signature: steal high, guest's own utilisation unchanged.

# --- 3. Bottleneck B: memory pressure. --------------------------------
for n in 2 3 4; do
  G "n$n" "nohup stress-ng --vm 1 --vm-bytes 85% --vm-keep \
           --timeout 90s >/dev/null 2>&1 &"
done
sleep 20
echo "=== B: memory pressure ==="
sudo virsh dommemstat vss1 | grep -E 'balloon|swap|available'
vmstat 2 3 | tail -2                    # host si/so columns
free -m | head -2
# The signature: balloon moving, or host swap non-zero. Different fix.

# --- 4. Bottleneck C: storage queueing. -------------------------------
pkill stress-ng 2>/dev/null || true ; sleep 5
G vss1 "nohup fio --name=load --rw=randwrite --bs=4k --size=512M \
        --numjobs=4 --time_based --runtime=60 --direct=1 \
        >/dev/null 2>&1 &"
sleep 10
echo "=== C: storage queueing ==="
iostat -x 2 3 | awk '/Device|vd|sd|nvme/' | tail -6
# The signature: aqu-sz (queue depth) and await (latency) both high,
# while %util may already have been 100% harmlessly before. Utilisation
# against saturation, section 18.2.

# --- 5. Percentiles, not means. ---------------------------------------
G vss1 "fio --name=lat --rw=randread --bs=4k --size=256M --direct=1 \
        --runtime=30 --time_based --output-format=json" \
  | grep -E '"mean"|"percentile"|"95.000000"|"99.000000"' | head -8
# Compare the mean with p95 and p99 on the same run.

# --- 6. Tear down. ----------------------------------------------------
for n in 2 3 4; do
  sudo virsh destroy "n$n" 2>/dev/null || true
  sudo virsh undefine "n$n" --remove-all-storage 2>/dev/null || true
done
sudo virsh destroy vss1 || true
