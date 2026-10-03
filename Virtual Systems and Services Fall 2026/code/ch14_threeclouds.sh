#!/usr/bin/env bash
# ch14_threeclouds.sh -- one guest on three providers, then delete.
set -euo pipefail

# --- 1. Authenticate to whichever accounts you have. -------------------
aws sts get-caller-identity   2>/dev/null || echo "no aws credentials"
az account show --query name  2>/dev/null || echo "no azure login"
gcloud config get-value project 2>/dev/null || echo "no gcp project"

# --- 2. AWS: a free-tier instance. ------------------------------------
AMI=$(aws ssm get-parameters --names \
  /aws/service/ami-amazon-linux-latest/al2023-ami-kernel-6.1-x86_64 \
  --query 'Parameters[0].Value' --output text)
aws ec2 run-instances --image-id "$AMI" --instance-type t3.micro \
  --count 1 --tag-specifications \
  'ResourceType=instance,Tags=[{Key=Name,Value=vss-lab}]' \
  --query 'Instances[0].InstanceId' --output text | tee aws-id.txt

# --- 3. Azure: the same shape. ----------------------------------------
az group create -n vss-lab -l eastus -o none
az vm create -g vss-lab -n vss1 --image Ubuntu2204 --size Standard_B1s \
  --generate-ssh-keys --public-ip-sku Basic -o table

# --- 4. GCP: likewise. ------------------------------------------------
gcloud compute instances create vss1 --machine-type=e2-micro \
  --zone=us-central1-a --image-family=debian-12 \
  --image-project=debian-cloud

# --- 5. Ask each guest what it is running on. -------------------------
# Every provider exposes a metadata service on 169.254.169.254. The
# answers show you the hypervisor of Chapter 5 underneath all three.
cat <<'PROBE' > probe.sh
systemd-detect-virt
sudo dmidecode -s system-manufacturer 2>/dev/null
curl -s -H 'Metadata-Flavor: Google' \
  169.254.169.254/computeMetadata/v1/instance/machine-type 2>/dev/null
curl -s 169.254.169.254/latest/meta-data/instance-type 2>/dev/null
PROBE
echo "copy probe.sh to each instance and run it"

# --- 6. What each one costs, from the provider's own API. -------------
aws pricing get-products --service-code AmazonEC2 --region us-east-1 \
  --filters 'Type=TERM_MATCH,Field=instanceType,Value=m7i.xlarge' \
            'Type=TERM_MATCH,Field=location,Value=US East (N. Virginia)' \
  --max-results 1 --query 'PriceList[0]' --output text 2>/dev/null \
  | head -c 400 || echo "needs the pricing:GetProducts permission"

# --- 7. TEAR DOWN. Do this even if a step above failed. ---------------
aws ec2 terminate-instances --instance-ids "$(cat aws-id.txt)" \
  --query 'TerminatingInstances[0].CurrentState.Name' --output text || true
az group delete -n vss-lab --yes --no-wait || true
gcloud compute instances delete vss1 --zone=us-central1-a --quiet || true

# --- 8. Confirm nothing survives. -------------------------------------
aws ec2 describe-instances --filters 'Name=tag:Name,Values=vss-lab' \
  'Name=instance-state-name,Values=running,pending' \
  --query 'Reservations[].Instances[].InstanceId' --output text
az group list --query "[?name=='vss-lab'].name" -o tsv
gcloud compute instances list --filter='name=vss1' --format='value(name)'
# All three should print nothing.
