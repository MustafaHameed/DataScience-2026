#!/usr/bin/env bash
# ch20_gpu.sh -- passthrough, MIG and what utilisation really reports.
set -euo pipefail

# --- 1. Is there an accelerator here at all? ---------------------------
if ! command -v nvidia-smi >/dev/null; then
  echo "no GPU on this host -- read code/ch20_capture.txt instead"
  sed -n '1,40p' "$(dirname "$0")/ch20_capture.txt" 2>/dev/null || true
  exit 0
fi
nvidia-smi --query-gpu=name,memory.total,mig.mode.current \
  --format=csv

# --- 2. The IOMMU group: what passthrough would have to take. ---------
# Chapter 4: the whole group moves together, or nothing does.
for d in $(lspci -D | awk '/NVIDIA/{print $1}'); do
  grp=$(basename "$(dirname "$(readlink -f \
        "/sys/bus/pci/devices/$d/iommu_group")")")
  echo "device $d is in IOMMU group $grp"
  ls "/sys/kernel/iommu_groups/$grp/devices/"
done
# A card and its audio function usually share a group: both must go.

# --- 3. Enable MIG, if this card supports it. -------------------------
sudo nvidia-smi -i 0 -mig 1 || \
  echo "MIG not supported on this card -- see the alertbox"
sudo nvidia-smi mig -i 0 -lgip      # the fixed profiles available

# --- 4. Create four instances, and see them as separate devices. ------
sudo nvidia-smi mig -i 0 -cgi 2g.20gb,2g.20gb,2g.20gb,2g.20gb -C || true
nvidia-smi -L                        # four MIG UUIDs, one card

# --- 5. Run the same job on one instance, then on all four. -----------
JOB='python3 -c "
import time
try:
    import torch
except ImportError:
    raise SystemExit(\"install torch to run this step\")
a = torch.randn(8192, 8192, device=\"cuda\")
t = time.time()
for _ in range(50): a = a @ a.T / 1e4
torch.cuda.synchronize(); print(f\"{time.time()-t:.1f} s\")"'

mig_uuids () {
  nvidia-smi -L | awk -F'UUID: ' '/MIG/{print $2}' | tr -d ')'
}
UUID=$(mig_uuids | head -1)
echo "=== alone on one instance ==="
CUDA_VISIBLE_DEVICES="$UUID" bash -c "$JOB"

echo "=== all four instances busy together ==="
mig_uuids | while read -r u; do
  CUDA_VISIBLE_DEVICES="$u" bash -c "$JOB" &
done
wait
# Under MIG the single-instance time and the all-busy time should be
# the SAME. That is the guarantee the worked example is about.

# --- 6. What "utilisation" actually reports. --------------------------
nvidia-smi --query-gpu=utilization.gpu,utilization.memory,\
memory.used,memory.total --format=csv -l 1 -c 5
# utilization.gpu is "a kernel was resident during the sample", not
# "the hardware was fully used". See the pitfalls.

# --- 7. Tear down the partitions. -------------------------------------
sudo nvidia-smi mig -i 0 -dci || true
sudo nvidia-smi mig -i 0 -dgi || true
sudo nvidia-smi -i 0 -mig 0  || true
