#!/usr/bin/env bash
# ch07_devices.sh -- emulation, virtio and vhost, measured.
set -euo pipefail
ip_of () {
  sudo virsh domifaddr "$1" | awk '/ipv4/{split($4,a,"/"); print a[1]}'
}
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$(ip_of vss1)" "$1"; }

# --- 1. What device model is the guest using right now? ----------------
sudo virsh dumpxml vss1 | sed -n '/<interface/,/<\/interface>/p'
# <model type='virtio'/> or type='e1000'. This single line is the chapter.

# --- 2. Set up the measurement: iperf3 on the host, client in guest. ---
sudo apt-get install -y iperf3 sysstat
iperf3 -s -D                       # server on the host, port 5201
sudo virsh start vss1 || true ; sleep 25
G "sudo apt-get -qq install -y iperf3"

# --- 3. Emulated e1000: a software imitation of a 1999 card. ----------
sudo virsh destroy vss1
sudo virt-xml vss1 --edit --network model=e1000
sudo virsh start vss1 ; sleep 25
mpstat 1 12 > host-e1000.txt &     # host CPU while the test runs
G "iperf3 -c 192.168.122.1 -t 10 -P 4" | tail -4
wait ; tail -2 host-e1000.txt

# --- 4. virtio: the split-driver model of section 7.2. ----------------
sudo virsh destroy vss1
sudo virt-xml vss1 --edit --network model=virtio
sudo virsh start vss1 ; sleep 25
mpstat 1 12 > host-virtio.txt &
G "iperf3 -c 192.168.122.1 -t 10 -P 4" | tail -4
wait ; tail -2 host-virtio.txt

# --- 5. virtio + vhost: back-end moved into the host kernel. ----------
# vhost_net is usually already on; confirm, and confirm the guest uses it.
lsmod | grep vhost_net || sudo modprobe vhost_net
sudo virsh destroy vss1
sudo virt-xml vss1 --edit --network driver.name=vhost
sudo virsh start vss1 ; sleep 25
mpstat 1 12 > host-vhost.txt &
G "iperf3 -c 192.168.122.1 -t 10 -P 4" | tail -4
wait ; tail -2 host-vhost.txt

# --- 6. Look at the virtqueues themselves. ----------------------------
# One queue pair per vCPU is the usual arrangement; each is a ring in
# memory that both sides can see, exactly as the definition box says.
G "ls /sys/class/net/*/queues/ ; ethtool -l ens3 2>/dev/null || true"

# --- 7. SR-IOV, if the host has a server NIC that supports it. --------
# Read-only: creating VFs needs a capable card and IOMMU enabled.
for d in /sys/class/net/*/device/sriov_totalvfs; do
  [ -e "$d" ] && echo "$d -> $(cat "$d") VFs available"
done
ls /sys/kernel/iommu_groups 2>/dev/null | wc -l

# --- 8. Restore and tear down. ----------------------------------------
sudo virsh destroy vss1 || true
pkill iperf3 || true
