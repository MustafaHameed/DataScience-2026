#!/usr/bin/env bash
# ch17_migrate.sh -- pre-copy convergence, observed and then broken.
set -euo pipefail
DEST=${1:?usage: ch17_migrate.sh <destination-host>}

# --- 1. Both hosts must see the same disk. ----------------------------
# Shared storage is the precondition for everything in this chapter.
showmount -e "$DEST" 2>/dev/null || \
  echo "set up NFS or iSCSI first -- see Appendix A"
sudo virsh pool-list --all

# --- 2. Start the guest and give it something to serve. --------------
sudo virsh start vss1 || true ; sleep 25
IP=$(sudo virsh domifaddr vss1 | awk '/ipv4/{split($4,a,"/");print a[1]}')
G () { ssh -o StrictHostKeyChecking=no "ubuntu@$IP" "$1"; }
G "sudo apt-get -qq install -y stress-ng python3 && \
   nohup python3 -m http.server 8000 >/dev/null 2>&1 &"

# --- 3. Watch continuity: ping and HTTP, both running throughout. ----
ping -i 0.2 "$IP" > /tmp/ping.log &
PINGPID=$!
( while true; do curl -s -o /dev/null -w '%{http_code} ' \
    "http://$IP:8000/" ; sleep 0.2; done ) > /tmp/http.log &
CURLPID=$!

# --- 4. Migrate it, and time the whole thing. ------------------------
time sudo virsh migrate --live --verbose --persistent --undefinesource \
  vss1 "qemu+ssh://$DEST/system"

sleep 2 ; kill $PINGPID $CURLPID 2>/dev/null || true
echo "--- dropped pings ---"
grep -c 'icmp_seq' /tmp/ping.log
tail -2 /tmp/ping.log           # look at the loss percentage
echo "--- non-200 responses ---"
tr ' ' '\n' < /tmp/http.log | grep -vc '^200$' || echo 0
# A correct migration drops nothing. The 50 ms pause is inside one
# ping interval.

# --- 5. Now make it refuse to converge. ------------------------------
# stress-ng --vm writes continuously, so the dirty rate exceeds the link.
sudo virsh migrate --live vss1 "qemu+ssh://$DEST/system" || true   # back
sleep 10
G "nohup stress-ng --vm 2 --vm-bytes 75% --vm-keep --timeout 300s \
   >/dev/null 2>&1 &"
sleep 5
echo "--- migrating a guest that dirties faster than the link ---"
sudo virsh migrate --live --verbose --timeout 60 \
  --timeout-suspend vss1 "qemu+ssh://$DEST/system" || \
  echo "did not converge within 60 s -- exactly section 17.1"

# --- 6. Make it converge by throttling the guest on purpose. ---------
sudo virsh migrate --live --verbose --auto-converge \
  --auto-converge-initial 20 --auto-converge-increment 10 \
  vss1 "qemu+ssh://$DEST/system"
# The guest was deliberately slowed so that it could move.

# --- 7. Measure the dirty rate directly, for the arithmetic. ---------
sudo virsh domdirtyrate-calc vss1 --seconds 10 || true
sleep 12 ; sudo virsh domstats vss1 --dirtyrate | grep -i dirty

# --- 8. Unplanned failure: what HA does instead. ---------------------
# No graceful anything: cut the guest off and see what survives.
sudo virsh destroy vss1
echo "the guest is gone; HA would now restart it elsewhere, and"
echo "everything in its memory is lost. That is the difference."

# --- 9. Tear down. ---------------------------------------------------
sudo virsh start vss1 || true
