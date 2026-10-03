#!/usr/bin/env bash
# ch19_harden.sh -- what --privileged means, and what hardening removes.
set -euo pipefail

# --- 1. A default container: how much is still exposed? ----------------
docker run --rm alpine:3.20 id            # root, by default
docker run --rm alpine:3.20 sh -c \
  'cat /proc/self/status | grep -E "^Cap(Eff|Prm)"'
# Decode those with capsh --decode=<hex>; the default set is ~14 caps.

# --- 2. The kernel is shared, and the container can see it. -----------
docker run --rm alpine:3.20 uname -a      # the HOST kernel version
uname -a                                  # identical. Chapter 8.

# --- 3. --privileged: the host, with extra steps. DISPOSABLE HOST ONLY.
docker run --rm --privileged alpine:3.20 sh -c \
  'ls /dev/ | head -20; echo "--- host disks visible above ---"'
docker run --rm --privileged --pid=host alpine:3.20 sh -c \
  'ps aux | head -5'                      # the host's process table
# With --privileged you can mount the host root and chroot into it.
# That is the whole "escape": there was never a boundary to cross.

# --- 4. The Docker socket is equivalent. ------------------------------
docker run --rm -v /var/run/docker.sock:/var/run/docker.sock \
  docker:cli sh -c 'docker ps | head -3'
# Anything that can reach the socket can start a privileged container
# with / mounted. Never grant it to something you do not fully trust.

# --- 5. Now harden, measure by measure. -------------------------------
docker run --rm --user 10001:10001 alpine:3.20 id
docker run --rm --cap-drop=ALL alpine:3.20 sh -c \
  'cat /proc/self/status | grep ^CapEff'  # 0000000000000000
docker run --rm --read-only alpine:3.20 sh -c 'touch /x' 2>&1 | tail -1
docker run --rm --security-opt=no-new-privileges alpine:3.20 \
  sh -c 'echo no-new-privileges set'

# --- 6. seccomp: count the system calls actually reachable. -----------
docker run --rm --security-opt seccomp=unconfined alpine:3.20 \
  sh -c 'echo unconfined: every syscall the kernel has'
# The default profile already blocks ~44 of them, including some that
# have historically been escape primitives.
docker info --format '{{.SecurityOptions}}'

# --- 7. All of it together: what a hardened run looks like. -----------
docker run --rm \
  --user 10001:10001 \
  --cap-drop=ALL \
  --security-opt=no-new-privileges \
  --read-only --tmpfs /tmp \
  --pids-limit 100 \
  --memory 256m --cpus 0.5 \
  --network none \
  alpine:3.20 sh -c 'id; echo "hardened"; touch /tmp/ok && echo "tmp ok"'

# --- 8. And the alternative, when none of that is enough. -------------
# Under Kubernetes this is one field (Chapter 11):
cat <<'YAML'
apiVersion: v1
kind: Pod
metadata:
  name: untrusted
spec:
  runtimeClassName: kata        # a hardware boundary, per workload
  containers:
    - name: job
      image: registry.example/builder@sha256:9b8c...
      securityContext:
        runAsNonRoot: true
        runAsUser: 10001
        allowPrivilegeEscalation: false
        readOnlyRootFilesystem: true
        capabilities: { drop: ["ALL"] }
YAML

# --- 9. Inventory your own estate for the two fatal flags. ------------
docker ps -q | while read -r c; do
  priv=$(docker inspect -f '{{.HostConfig.Privileged}}' "$c")
  sock=$(docker inspect -f '{{range .Mounts}}{{.Source}} {{end}}' "$c" \
         | grep -c docker.sock || true)
  [ "$priv" = "true" ] && echo "PRIVILEGED: $c"
  [ "$sock" != "0" ] && echo "DOCKER SOCKET MOUNTED: $c"
done
echo "inventory complete"
