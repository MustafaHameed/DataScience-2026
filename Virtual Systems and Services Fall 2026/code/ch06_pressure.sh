#!/usr/bin/env bash
# ch06_pressure.sh -- ready time and the reclamation ladder, observed.
set -euo pipefail
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$1" "$2"; }

# --- 1. How many physical cores are we actually working with? ----------
nproc ; lscpu | grep -E 'Socket|Core|NUMA'

# --- 2. Build four guests, each with as many vCPUs as the host has. ----
# Deliberate oversubscription: 4 guests x nproc vCPUs on nproc cores.
for n in 1 2 3 4; do
  sudo qemu-img create -f qcow2 -b /var/lib/libvirt/images/vss1.qcow2 \
       -F qcow2 "/var/lib/libvirt/images/p$n.qcow2" 10G
  sudo virt-install --name "p$n" --memory 2048 --vcpus "$(nproc)" \
    --disk "/var/lib/libvirt/images/p$n.qcow2",format=qcow2 \
    --disk seed.iso,device=cdrom --os-variant ubuntu22.04 \
    --import --graphics none --noautoconsole
done
sleep 40

# --- 3. Baseline: one guest busy, three idle. --------------------------
ip_of () {
  sudo virsh domifaddr "$1" | awk '/ipv4/{split($4,a,"/"); print a[1]}'
}
IP1=$(ip_of p1)
G "$IP1" "sudo apt-get -qq install -y stress-ng sysstat"
G "$IP1" "stress-ng --cpu \$(nproc) --timeout 30s --metrics-brief"

# --- 4. Now make all four busy and read steal time in the guest. -------
# %st in the guest IS the ready time of section 6.1, seen from inside.
for n in 2 3 4; do
  IP=$(ip_of "p$n")
  G "$IP" "sudo apt-get -qq install -y stress-ng" || true
  G "$IP" "nohup stress-ng --cpu \$(nproc) --timeout 60s >/dev/null 2>&1 &"
done
sleep 5
G "$IP1" "mpstat 2 5 | tail -4"       # watch the %steal column climb

# --- 5. Memory pressure: ask for far more than the host has. -----------
# Watch the balloon move before anything swaps.
sudo virsh dommemstat p1
for n in 1 2 3 4; do
  G "$(ip_of "p$n")" "nohup stress-ng --vm 1 --vm-bytes 85% \
           --timeout 60s >/dev/null 2>&1 &" || true
done
sleep 20
sudo virsh dommemstat p1               # balloon / swap_in / swap_out
free -m ; vmstat 2 4                   # the host's own view

# --- 6. Right-size, and measure the same thing again. ------------------
# Two vCPUs each instead of nproc: the overcommit ratio drops fourfold.
for n in 1 2 3 4; do sudo virsh destroy "p$n"; \
  sudo virsh setvcpus "p$n" 2 --config --maximum --config; \
  sudo virsh setvcpus "p$n" 2 --config; sudo virsh start "p$n"; done
sleep 40
G "$IP1" "mpstat 2 5 | tail -4"        # %steal should be far lower

# --- 7. Tear down. -----------------------------------------------------
for n in 1 2 3 4; do
  sudo virsh destroy "p$n" || true
  sudo virsh undefine "p$n" --remove-all-storage || true
done
