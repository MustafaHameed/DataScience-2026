# Virtual Systems and Services — Fall 2026

Weekly lecture schedule for the BSIT elective **Virtual Systems and
Services**, Department of Information Technology, The Islamia University of
Bahawalpur.

## Where this course comes from

HEC published a full outline for this course exactly once: **HEC 2017**
(*Curriculum of BS & MS CS, SE & IT*, pp. 71–130) — 3 credit hours,
prerequisite *Programming Fundamentals*, a hypervisor course covering VMware,
Xen and Hyper-V, para-virtualized components and processor support for
virtualization. HEC 2023 keeps the title as an unspecified BSIT domain
elective (credit count only, no published contents). **HEC 2025 drops the
course entirely** — its ground is split between *Cloud Computing* (a new
major-core course, Semester VI) and *Microservices Architecture and Docker
Containers* (Network Infrastructure & Cloud Computing specialization). See
[`../HEC AI-DS Curriculum Map/`](../HEC%20AI-DS%20Curriculum%20Map/) for the
full edition-by-edition trace.

This schedule modernises the 2017 outline rather than reviving it verbatim:
the hypervisor foundations (Weeks 1–6) are taught close to the original
specification, then carried forward (Weeks 7–16) into containers, Kubernetes,
cloud service models and confidential computing — the ground HEC 2025 hands to
*Cloud Computing* and *Microservices Architecture and Docker Containers*. The
2017 set texts (Stanney's *Handbook of Virtual Environments*, Burdea's
*Virtual Reality Technology*) are virtual-reality references, not
virtualization references, and are not carried forward — see the schedule's
own footnote.

**No full handout exists yet for this course** — only the one-page weekly
schedule below.

## Build

```powershell
pdflatex -interaction=nonstopmode -file-line-error VSS_Fall2026_Lecture_Schedule.tex
```

One pass is enough (no table of contents, no cross-references, no
bibliography). Requires the same MiKTeX packages as the other course
schedules in this repository: `fontawesome5`, `ragged2e`, `microtype`,
`xcolor` (`table` option), `array`.

## Layout

```
VSS_Fall2026_Lecture_Schedule.tex   the one-page schedule (self-contained, own preamble)
```

Follows the same one-page schedule pattern as the other Fall 2026 courses
([`../Data Science Fall 2026/`](../Data%20Science%20Fall%202026/),
[`../Machine Learning Fall 2026/`](../Machine%20Learning%20Fall%202026/),
[`../Advanced Research Methodology Fall 2026/`](../Advanced%20Research%20Methodology%20Fall%202026/),
[`../CCNA/`](../CCNA/)): 16 teaching weeks, 6 colour-coded parts, a 3-tile
assessment-weight footer (Sessional 20% / Mid-term 30% / Final-term 50%).

## Structure

**Part I — Foundations of Virtualization** (1 What Virtualization Is · 2
Virtualization Architectures · 3 The Virtual Machine Monitor)

**Part II — Hypervisors and Compute Virtualization** (4 Hardware Support for
Virtualization · 5 Major Hypervisor Platforms · 6 CPU and Memory
Virtualization · 7 Paravirtualization and Device Virtualization)

**Part III — Containers and Lightweight Virtualization** (8 From Virtual
Machines to Containers · 9 Docker and the Container Ecosystem · 10 Container
Orchestration with Kubernetes · 11 Lightweight and Secure Virtualization)

**Part IV — Storage, Network and Cloud Virtualization** (12 Storage
Virtualization · 13 Network Virtualization and SDN · 14 Cloud Service and
Deployment Models · 15 Desktop and Application Virtualization)

**Part V — Management, Performance and Security** (16 Virtualization
Management Platforms · 17 High Availability, Live Migration and DR · 18
Performance Monitoring and Capacity Planning · 19 Security in Virtualized
Environments)

**Part VI — Emerging Hardware and Practice** (20 GPU and Accelerator
Virtualization · 21 Confidential Computing and Emerging Hardware · 22 Case
Studies in Enterprise and Cloud Virtualization · 23 Semester Project
Presentations)
