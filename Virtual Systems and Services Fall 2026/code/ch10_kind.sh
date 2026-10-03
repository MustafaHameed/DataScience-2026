#!/usr/bin/env bash
# ch10_kind.sh -- reconciliation, scheduling and probes, observed.
set -euo pipefail

# --- 1. A three-node cluster in containers on this machine. ------------
kind create cluster --name vss --config - <<'YAML'
kind: Cluster
apiVersion: kind.x-k8s.io/v1alpha4
nodes:
  - role: control-plane
  - role: worker
  - role: worker
YAML
kubectl get nodes -o wide
# Each "node" is a container (Chapter 9) running a kubelet. Containers
# all the way down.

# --- 2. Declare a desired state and watch six loops reach it. ----------
kubectl create deployment web --image=nginx:1.27 --replicas=3
kubectl get deploy,rs,pods -o wide
kubectl get events --sort-by=.lastTimestamp | tail -12
# Read the events bottom-up: scheduled, pulling, created, started.

# --- 3. Reconciliation: delete a Pod and time the replacement. ---------
kubectl get pods -l app=web
kubectl delete pod -l app=web --field-selector status.phase=Running \
  --wait=false | head -1
sleep 3 ; kubectl get pods -l app=web
# Nothing "handled" the deletion. Two is not three, so a loop acted.

# --- 4. A Service: a stable address in front of changing Pods. --------
kubectl expose deployment web --port=80 --target-port=80
kubectl get svc web
kubectl get endpoints web          # the Pod IPs, maintained by a controller
kubectl delete pod -l app=web --wait=false | head -1
sleep 5 ; kubectl get endpoints web   # different Pod IPs, same Service IP

# --- 5. Make a Pod unschedulable, and read why. -----------------------
kubectl create deployment big --image=nginx:1.27
kubectl set resources deployment big --requests=cpu=8,memory=32Gi
kubectl get pods -l app=big
kubectl describe pod -l app=big | sed -n '/Events/,$p' | tail -6
# "0/3 nodes are available: Insufficient cpu" -- the worked example,
# reported by the scheduler itself.

# --- 6. QoS classes, derived from what you declared. ------------------
kubectl run guaranteed --image=nginx:1.27 \
  --overrides='{"spec":{"containers":[{"name":"c","image":"nginx:1.27",
  "resources":{"requests":{"cpu":"100m","memory":"64Mi"},
  "limits":{"cpu":"100m","memory":"64Mi"}}}]}}'
kubectl run besteffort --image=nginx:1.27
sleep 5
kubectl get pod guaranteed besteffort \
  -o custom-columns=NAME:.metadata.name,QOS:.status.qosClass

# --- 7. The cascading-restart outage, on purpose. ---------------------
# A liveness probe pointing at something that will fail.
kubectl create deployment fragile --image=nginx:1.27 --replicas=3
kubectl patch deployment fragile --type=json -p='[{"op":"add",
  "path":"/spec/template/spec/containers/0/livenessProbe",
  "value":{"httpGet":{"path":"/nope","port":80},"periodSeconds":5}}]'
sleep 45
kubectl get pods -l app=fragile      # RESTARTS climbing on every replica
kubectl describe pod -l app=fragile | grep -A2 'Liveness' | head -6
# Now imagine /nope were a real endpoint that checks a slow database.

# --- 8. Tear down. ----------------------------------------------------
kind delete cluster --name vss
