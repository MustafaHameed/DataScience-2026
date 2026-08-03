# Lab code

Runnable extracts from the **Try It Yourself** boxes in the handout. Each file is
named for the chapter it comes from.

Files with no external data dependency run standalone and are listed as
*self-contained* below — those are the ones used to verify the labs. The rest
expect a dataset you supply; the handout box says which.

| File | Chapter | Self-contained? |
|---|---|---|
| `ch04_distance_descent.py` | 4 — The Mathematical Toolkit | yes |
| `ch06_base_rate.py` | 6 — Probability and Uncertainty | yes |
| `ch10_confounder.py` | 10 — Relationships and Causality | yes |
| `ch17_network_from_scratch.py` | 17 — Neural Network Foundations | yes |
| `ch28_fairness.py` | 28 — Responsible AI | yes |
| `ch29_provenance.py` | 29 — Research Methods | yes (needs a git repo) |

## Running

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install numpy pandas scikit-learn
python ch06_base_rate.py
```
