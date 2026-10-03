#!/usr/bin/env bash
# ch15_vdi.sh -- a remote desktop, measured, then made to feel slow.
set -euo pipefail

# --- 1. A desktop guest, from the Chapter 1 image. ---------------------
sudo virsh start vss1 || true ; sleep 25
ip_of () {
  sudo virsh domifaddr "$1" | awk '/ipv4/{split($4,a,"/"); print a[1]}'
}
IP=$(ip_of vss1)
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$IP" "$1"; }
G "sudo apt-get -qq install -y xfce4 xrdp firefox && sudo systemctl \
   enable --now xrdp"

# --- 2. Connect with an RDP client and leave it idle. -----------------
# From the host: xfreerdp /v:$IP /u:ubuntu, or any RDP client.
sudo apt-get install -y freerdp2-x11 iftop tcpdump
echo "connect now, then press return"; read -r _

# --- 3. Measure the three cases of the figure. ------------------------
measure () {
  echo "=== $1 ==="
  sudo timeout 20 tcpdump -ni any "host $IP and port 3389" -q 2>/dev/null \
    | wc -l
}
measure "idle -- do nothing for 20 s"
echo "now type continuously in a text editor, then press return"; read -r _
measure "typing"
echo "now play a video full screen, then press return"; read -r _
measure "video"
# Packet counts stand in for bandwidth; the ratios are what matter.

# --- 4. Now the thing users actually feel: latency. -------------------
# Add 150 ms each way and use the session again. Nothing else changes.
sudo tc qdisc add dev virbr0 root netem delay 150ms
echo "use the session now -- type a sentence. Press return when done."
read -r _
sudo tc qdisc del dev virbr0 root
echo "and again with no added delay. The bandwidth was identical."

# --- 5. The boot storm, at small scale. -------------------------------
# Clone four desktops from one base and start them together; watch IOPS.
for n in 1 2 3 4; do
  sudo qemu-img create -f qcow2 -b /var/lib/libvirt/images/vss1.qcow2 \
       -F qcow2 "/var/lib/libvirt/images/d$n.qcow2" 20G
done
ls -lh /var/lib/libvirt/images/d*.qcow2   # linked clones: nearly empty
iostat -x 2 2 > /tmp/before.txt
for n in 1 2 3 4; do
  sudo virt-install --name "d$n" --memory 1024 --vcpus 1 \
    --disk "/var/lib/libvirt/images/d$n.qcow2",format=qcow2 \
    --import --graphics none --noautoconsole --os-variant ubuntu22.04 &
done
wait ; iostat -x 5 3 | tail -12      # the storm, four desktops wide

# --- 6. Tear down. ----------------------------------------------------
for n in 1 2 3 4; do
  sudo virsh destroy "d$n" 2>/dev/null || true
  sudo virsh undefine "d$n" --remove-all-storage 2>/dev/null || true
done
sudo virsh destroy vss1 || true
