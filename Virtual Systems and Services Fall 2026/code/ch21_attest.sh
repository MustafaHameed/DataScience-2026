#!/usr/bin/env bash
# ch21_attest.sh -- is this guest confidential, and can it prove it?
set -euo pipefail

# --- 1. Does this processor support any of the three? -----------------
grep -o -m1 -E '\<(sev|sev_es|sev_snp|tdx_guest)\>' /proc/cpuinfo || \
  echo "no TEE features on this host"
ls /sys/firmware/ 2>/dev/null | grep -iE 'tdx|sev' || true
dmesg | grep -iE 'SEV-SNP|TDX|Memory Encryption' | head -5 || true

# --- 2. From inside a guest: am I confidential? -----------------------
# Each platform exposes itself differently; check all three.
[ -e /dev/sev-guest ]  && echo "AMD SEV-SNP guest device present"
[ -e /dev/tdx_guest ]  && echo "Intel TDX guest device present"
systemd-detect-virt
# On a normal guest, none of these exist. That is the point of step 4.

# --- 3. Ask the hardware for an attestation report. -------------------
# snpguest (AMD) or trustee/tdx tools (Intel). A 64-byte nonce from the
# verifier goes in; a signed report comes out.
NONCE=$(openssl rand -hex 32)
if [ -e /dev/sev-guest ] && command -v snpguest >/dev/null; then
  snpguest report report.bin request.txt --random
  snpguest display report report.bin | head -30
else
  echo "no TEE here -- using the captured report in code/"
  cp "$(dirname "$0")/ch21_report.txt" report.txt 2>/dev/null || true
  sed -n '1,30p' report.txt 2>/dev/null || true
fi

# --- 4. The three fields that matter, and why. ------------------------
cat <<'NOTES'
  measurement    the hash of the guest's initial state. Compare this
                 with what you built; if it differs, something else
                 is running and you must NOT release the key.
  report_data    your nonce, echoed back. Proves the report is fresh
                 and not a recording of an earlier genuine boot.
  signature +    chains to a key the silicon vendor certified. Proves
  cert chain     the report came from real hardware, not from software
                 pretending very convincingly.
NOTES

# --- 5. Verify the chain. This is the step people skip. ---------------
if command -v snpguest >/dev/null; then
  snpguest fetch ca-der milan ./certs           # vendor certificates
  snpguest fetch vcek-der milan ./certs report.bin
  snpguest verify certs ./certs
  snpguest verify attestation ./certs report.bin && \
    echo "VERIFIED: genuine hardware, measurement as reported"
else
  echo "verification exercise: compare the measurement in report.txt"
  echo "with code/ch21_expected_measurement.txt and decide."
fi

# --- 6. Now the attack this defends against. --------------------------
# A hostile operator boots a DIFFERENT image that asks for your key.
# Its memory is perfectly encrypted while it exfiltrates. The only
# thing that stops it is this comparison.
D=$(dirname "$0")
hex () { grep -o -m1 -E '[0-9a-f]{96}' "$1"; }   # the 48-byte measurement
GOT=$(hex report.txt)
WANT=$(hex "$D/ch21_expected_measurement.txt")
echo "got:  ${GOT:-<none>}"
echo "want: ${WANT:-<none>}"
if [ -n "$GOT" ] && [ "$GOT" = "$WANT" ]; then
  echo "MATCH -- release the key"
else
  echo "DIFFERS -- refuse. This comparison is the whole mechanism."
fi

# --- 7. What it costs: the same benchmark, confidential or not. ------
# On a cloud that offers both, run this on a confidential and a normal
# instance of the same family and compare.
sysbench cpu --threads=4 --time=20 run 2>/dev/null | grep -E 'events/s' \
  || echo "install sysbench to measure the overhead"
dd if=/dev/zero of=/tmp/t bs=1M count=2048 oflag=direct 2>&1 | tail -1
rm -f /tmp/t
# I/O usually shows the larger penalty: encrypted guests bounce device
# transfers through shared buffers.
