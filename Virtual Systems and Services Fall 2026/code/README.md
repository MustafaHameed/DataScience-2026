# Lab code — Virtual Systems and Services, Fall 2026

One file per lab, named for its chapter. **These are extracted from the
handout, not written here**: each lab is authored once inside its `lab` box in
`../parts/`, and `sync_labs.py` writes it out. Editing a file in this directory
is pointless — the next sync overwrites it. Fix the chapter instead.

```bash
python sync_labs.py            # write the files from the handout
python sync_labs.py --check    # fail if anything has drifted
```

## Before you start

Appendix A of the handout installs and verifies everything these need, in one
pass. Do that first; every script below assumes it.

Three labs need hardware not everyone has. Each ships captured output here, and
the analysis — which is the actual exercise — is unchanged.

| Lab | Needs | Read instead |
| --- | --- | --- |
| `ch04_exits.sh` | nested virtualization | `ch04_exits.txt` |
| `ch20_gpu.sh` | an NVIDIA GPU with MIG | `ch20_capture.txt` |
| `ch21_attest.sh` | an SEV-SNP or TDX processor | `ch21_report.txt`, `ch21_expected_measurement.txt` |

## The files

| File | Chapter | What it does | Needs |
| --- | --- | --- | --- |
| `ch01_first_vm.sh` | 1 | Create, snapshot, break and restore a guest | KVM |
| `ch02_overhead.sh` | 2 | The cost of the layer, three ways | KVM, Docker |
| `ch03_popf.c` | 3 | `POPF` failing silently in user mode | gcc |
| `ch04_exits.sh` | 4 | VMX, EPT and the exit rate | KVM, `perf` |
| `ch05_platforms.sh` | 5 | One guest on two platforms | KVM, VirtualBox |
| `ch06_pressure.sh` | 6 | Ready time and the reclamation ladder | **a disposable host** |
| `ch07_devices.sh` | 7 | Emulated, virtio and vhost, measured | KVM, `iperf3` |
| `ch08_byhand.sh` | 8 | A container from kernel primitives, no runtime | root |
| `ch09_build.sh` | 9 | The build cache, measured, then fixed | Docker |
| `ch10_kind.sh` | 10 | A cluster on one laptop, and three failure modes | `kind` |
| `ch11_microvm.sh` | 11 | Booting a microVM and timing it | KVM |
| `ch12_storage.sh` | 12 | A thin pool walked off its cliff | LVM, root |
| `ch13_overlay.sh` | 13 | A VXLAN overlay, and an MTU black hole | root |
| `ch14_threeclouds.sh` | 14 | One workload on three clouds | **cloud accounts** |
| `ch15_vdi.sh` | 15 | What a remote desktop sends, and what users feel | KVM, RDP client |
| `ch16_iac.sh` | 16 | An image pipeline and declarative provisioning | Terraform |
| `ch17_migrate.sh` | 17 | Live migration, then made to fail | **two hosts** |
| `ch18_bottleneck.sh` | 18 | Three bottlenecks, three signatures | `sysstat`, `fio` |
| `ch19_harden.sh` | 19 | What `--privileged` means, and what hardening removes | **a disposable host** |
| `ch20_gpu.sh` | 20 | Partitioning an accelerator | a GPU |
| `ch21_attest.sh` | 21 | Attestation, verified | SEV-SNP or TDX |
| `ch23_proposal.md` | 23 | Your project proposal template | your team |

Chapter 22's lab is a written analysis with no code; the brief is in the
chapter.

## Read these warnings

- **`ch06_pressure.sh` and `ch19_harden.sh` will make a machine unusable.**
  The first drives a host into CPU and memory exhaustion deliberately; the
  second runs `--privileged` containers that can see and modify the host. Use a
  machine you do not care about.
- **`ch12_storage.sh` fills a thin pool until writes fail.** That is the point
  of it, and it is done on a loopback file rather than real storage.
- **`ch14_threeclouds.sh` creates billable resources.** Step 7 tears them down
  and step 8 verifies that nothing survived. Run both even if an earlier step
  failed, and set the budget alert from Appendix A first.
- Every script starts with `set -euo pipefail`, so it stops at the first
  failure rather than continuing with a bad assumption.

## State of verification

Honest accounting, because the labs are the part of a handout most likely to
rot:

- **All 20 shell labs pass `bash -n`** (syntax), and all listings are
  ASCII-only and within 76 columns — `../checklabs.py` enforces both.
- **`ch21_attest.sh`'s comparison logic has been run** against the shipped
  capture files and correctly reports a match.
- **The labs have not been executed end to end on the machine that produced
  this handout**, which has no hypervisor, no container engine and no C
  compiler. They are written against documented tool behaviour. Treat the first
  run of each in a given environment as part of the lab.

If a command has changed under you, the fix belongs in `../parts/`, not here.
