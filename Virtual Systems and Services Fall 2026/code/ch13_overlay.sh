#!/usr/bin/env bash
# ch13_overlay.sh -- a virtual switch, a VXLAN tunnel and an MTU black hole.
set -euo pipefail

# --- 1. A virtual switch, and two "guests" as namespaces. --------------
sudo ip link add br-vss type bridge
sudo ip link set br-vss up
for n in 1 2; do
  sudo ip netns add "g$n"
  sudo ip link add "veth$n" type veth peer name "in$n"
  sudo ip link set "veth$n" master br-vss
  sudo ip link set "veth$n" up
  sudo ip link set "in$n" netns "g$n"
  sudo ip netns exec "g$n" ip addr add "10.10.0.$n/24" dev "in$n"
  sudo ip netns exec "g$n" ip link set "in$n" up
  sudo ip netns exec "g$n" ip link set lo up
done
# This is container networking, in nine lines.
sudo ip netns exec g1 ping -c 2 10.10.0.2
bridge fdb show br br-vss | head -4     # the switch has learned MACs

# --- 2. Now the underlay: two more namespaces playing "hosts". --------
sudo ip netns add h1 ; sudo ip netns add h2
sudo ip link add u1 type veth peer name u2
sudo ip link set u1 netns h1 ; sudo ip link set u2 netns h2
sudo ip netns exec h1 ip addr add 192.168.50.1/24 dev u1
sudo ip netns exec h2 ip addr add 192.168.50.2/24 dev u2
for h in h1 h2; do
  sudo ip netns exec "$h" ip link set lo up
done
sudo ip netns exec h1 ip link set u1 up
sudo ip netns exec h2 ip link set u2 up
sudo ip netns exec h1 ping -c 2 192.168.50.2

# --- 3. A VXLAN tunnel across it: VNI 100. ----------------------------
sudo ip netns exec h1 ip link add vx0 type vxlan id 100 \
     remote 192.168.50.2 local 192.168.50.1 dstport 4789 dev u1
sudo ip netns exec h2 ip link add vx0 type vxlan id 100 \
     remote 192.168.50.1 local 192.168.50.2 dstport 4789 dev u2
sudo ip netns exec h1 ip addr add 172.16.0.1/24 dev vx0
sudo ip netns exec h2 ip addr add 172.16.0.2/24 dev vx0
sudo ip netns exec h1 ip link set vx0 up
sudo ip netns exec h2 ip link set vx0 up
sudo ip netns exec h1 ip -d link show vx0 | grep -o 'vxlan id 100'

# --- 4. Small packets work. ------------------------------------------
sudo ip netns exec h1 ping -c 2 172.16.0.2

# --- 5. Now the black hole: large packets, don't-fragment set. --------
# The underlay is 1500; VXLAN needs 50; so 1450 is the real limit.
sudo ip netns exec h1 ip link set vx0 mtu 1500    # deliberately wrong
echo "--- 1400 bytes: fits ---"
sudo ip netns exec h1 ping -c 2 -M do -s 1400 172.16.0.2 || true
echo "--- 1472 bytes: does not fit, and will not say so ---"
sudo ip netns exec h1 ping -c 2 -W 2 -M do -s 1472 172.16.0.2 || \
  echo "no reply -- exactly the worked example's symptom"

# --- 6. Fix 1: lower the overlay MTU to the arithmetic. ---------------
sudo ip netns exec h1 ip link set vx0 mtu 1450
sudo ip netns exec h2 ip link set vx0 mtu 1450
sudo ip netns exec h1 ping -c 2 -M do -s 1422 172.16.0.2

# --- 7. Fix 2: raise the underlay instead, and keep 1500 inside. ------
sudo ip netns exec h1 ip link set u1 mtu 9000
sudo ip netns exec h2 ip link set u2 mtu 9000
sudo ip netns exec h1 ip link set vx0 mtu 8950
sudo ip netns exec h2 ip link set vx0 mtu 8950
sudo ip netns exec h1 ping -c 2 -M do -s 8000 172.16.0.2

# --- 8. Watch the encapsulation on the wire. -------------------------
sudo timeout 6 ip netns exec h2 tcpdump -ni u2 -c 4 udp port 4789 &
sleep 1
sudo ip netns exec h1 ping -c 3 172.16.0.2 >/dev/null ; wait || true
# Each capture line is one outer packet carrying one whole inner frame.

# --- 9. Tear down. ---------------------------------------------------
for n in g1 g2 h1 h2; do sudo ip netns del "$n" 2>/dev/null || true; done
sudo ip link del br-vss 2>/dev/null || true
