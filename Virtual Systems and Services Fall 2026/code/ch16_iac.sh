#!/usr/bin/env bash
# ch16_iac.sh -- an image pipeline and declarative provisioning.
set -euo pipefail
mkdir -p ~/vsiac && cd ~/vsiac

# --- 1. The build stage: one patched image, made reproducibly. --------
# Packer would do this in a pipeline; here, the same steps by hand.
BASE=https://cloud-images.ubuntu.com/jammy/current
wget -nc -O base.img "$BASE/jammy-server-cloudimg-amd64.img"
cp base.img golden-2026-10.qcow2
sudo virt-customize -a golden-2026-10.qcow2 \
  --update --install qemu-guest-agent,chrony \
  --run-command 'systemctl enable qemu-guest-agent' \
  --truncate /etc/machine-id          # or every clone shares an identity
qemu-img info golden-2026-10.qcow2 | head -4

# --- 2. The publish stage: versioned, and never edited again. ---------
sudo mkdir -p /var/lib/libvirt/images/templates
sudo cp golden-2026-10.qcow2 /var/lib/libvirt/images/templates/
sudo chmod 444 /var/lib/libvirt/images/templates/golden-2026-10.qcow2
ls -l /var/lib/libvirt/images/templates/

# --- 3. Declare the estate. ------------------------------------------
cat > main.tf <<'TF'
terraform {
  required_providers {
    libvirt = { source = "dmacvicar/libvirt", version = "~> 0.8" }
  }
}
provider "libvirt" { uri = "qemu:///system" }

variable "guests" {
  type    = map(object({ memory = number, vcpu = number }))
  default = {
    web1 = { memory = 1024, vcpu = 1 }
    web2 = { memory = 1024, vcpu = 1 }
    db1  = { memory = 2048, vcpu = 2 }
  }
}

resource "libvirt_volume" "disk" {
  for_each       = var.guests
  name           = "${each.key}.qcow2"
  base_volume_id = "/var/lib/libvirt/images/templates/golden-2026-10.qcow2"
  size           = 10737418240
}

resource "libvirt_cloudinit_disk" "init" {
  for_each  = var.guests
  name      = "${each.key}-init.iso"
  user_data = <<EOT
#cloud-config
hostname: ${each.key}
password: vss2026
chpasswd: { expire: False }
ssh_pwauth: True
EOT
}

resource "libvirt_domain" "vm" {
  for_each  = var.guests
  name      = each.key
  memory    = each.value.memory
  vcpu      = each.value.vcpu
  cloudinit = libvirt_cloudinit_disk.init[each.key].id
  disk { volume_id = libvirt_volume.disk[each.key].id }
  network_interface { network_name = "default" }
  console { type = "pty", target_port = "0" }
}
TF

# --- 4. Plan first. The plan IS the review. --------------------------
terraform init -input=false
terraform plan -out=tfplan | tail -20

# --- 5. Apply, and then prove idempotence. ---------------------------
terraform apply -auto-approve tfplan
sudo virsh list --all
terraform plan | tail -3        # "No changes" -- this is the property
                                # a shell script of virsh does not have.

# --- 6. Reconfigure declaratively: change one number. ----------------
sed -i 's/db1  = { memory = 2048/db1  = { memory = 4096/' main.tf
terraform plan | grep -E 'memory|will be'
terraform apply -auto-approve

# --- 7. Create drift by hand, then watch the tool notice. ------------
sudo virsh setmaxmem web1 3145728 --config
sudo virsh setmem web1 3145728 --config
terraform plan | tail -8        # the difference, reported

# --- 8. Tear the whole estate down with one command. -----------------
terraform destroy -auto-approve
sudo virsh list --all
