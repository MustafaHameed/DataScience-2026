#!/usr/bin/env bash
# ch08_byhand.sh -- namespaces, cgroups and overlayfs, without a runtime.
set -euo pipefail
ROOT=/tmp/vsroot

# --- 1. What namespaces does an ordinary process already have? ---------
# Every process is in a namespace; a container just gets different ones.
ls -l /proc/self/ns/

# --- 2. A root filesystem to be the container's world. -----------------
sudo apt-get install -y debootstrap util-linux
sudo mkdir -p "$ROOT"
sudo debootstrap --variant=minbase jammy "$ROOT" \
     http://archive.ubuntu.com/ubuntu/

# --- 3. Enter new namespaces. This is the whole mechanism. -------------
# --pid --mount --uts --ipc --net: one flag per namespace from the table.
sudo unshare --pid --mount --uts --ipc --net --fork --mount-proc \
     chroot "$ROOT" /bin/bash -c '
       hostname container-1
       mount -t proc proc /proc
       echo "--- inside ---"
       hostname            # the UTS namespace
       ps aux              # the PID namespace: we are PID 1, host invisible
       ip addr             # the network namespace: only lo, and it is down
       exit'

# --- 4. Prove the host is unaffected. ----------------------------------
hostname ; ps aux | wc -l

# --- 5. Limit it with a control group. ---------------------------------
# cgroup v2: one unified tree under /sys/fs/cgroup.
sudo mkdir -p /sys/fs/cgroup/vsdemo
echo "+cpu +memory" | sudo tee /sys/fs/cgroup/cgroup.subtree_control
echo "50000 100000" | sudo tee /sys/fs/cgroup/vsdemo/cpu.max   # half a core
echo "100M"         | sudo tee /sys/fs/cgroup/vsdemo/memory.max

# Put a busy process in it and watch it be throttled, not killed.
sudo bash -c 'echo $$ > /sys/fs/cgroup/vsdemo/cgroup.procs; \
              timeout 10 yes > /dev/null' &
sleep 11
cat /sys/fs/cgroup/vsdemo/cpu.stat | grep throttled

# --- 6. Now cross the memory limit, and watch it be killed. -----------
# CPU throttles; memory kills. This is the asymmetry of section 8.3.
sudo bash -c 'echo $$ > /sys/fs/cgroup/vsdemo/cgroup.procs; \
              head -c 200M /dev/zero | tail -c 1' || \
  echo "killed, exit $? -- 137 is 128+9, the OOM killer"
dmesg | tail -3 | grep -i 'out of memory' || true

# --- 7. Share one root between two containers with overlayfs. ---------
sudo mkdir -p /tmp/ovl/{upper1,upper2,work1,work2,m1,m2}
for i in 1 2; do
  sudo mount -t overlay overlay \
    -o lowerdir="$ROOT",upperdir=/tmp/ovl/upper$i,workdir=/tmp/ovl/work$i \
    /tmp/ovl/m$i
done
# Write in one; the other and the shared lower layer are untouched.
sudo touch /tmp/ovl/m1/etc/only-in-one
ls /tmp/ovl/m1/etc/only-in-one
ls /tmp/ovl/m2/etc/only-in-one 2>&1 | tail -1
sudo du -sh /tmp/ovl/upper1 /tmp/ovl/upper2 "$ROOT"

# --- 8. Tear down. ----------------------------------------------------
for i in 1 2; do sudo umount /tmp/ovl/m$i || true; done
sudo rmdir /sys/fs/cgroup/vsdemo || true
sudo rm -rf /tmp/ovl "$ROOT"
