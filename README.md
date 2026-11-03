# PhD Workplan — Geoff Whittle-Walls

Python script that generates a fully-formatted Excel workplan for a
**PhD (Direct Entry) · Thesis Incorporating Publications**
at the University of Essex (supervisors: Marie Juanchich & Miro Sirota).

## What it builds

Running the script produces `PhD_Workplan_GWW.xlsx` with five tabs:

| Tab | Contents |
|-----|----------|
| **Workplan** | 150 tasks across 14 phases — planned/actual dates, status, % done, flags |
| **Dashboard** | Summary metrics and phase gate traffic lights |
| **Quality Gates** | Phase-by-phase go/no-go criteria linked to supervisory meetings |
| **Deliverables** | Full outputs register (papers, chapters, thesis milestones) |

## Quick start

```bash
pip install -r requirements.txt
python build_workplan.py
```

The file `PhD_Workplan_GWW.xlsx` is written to the current directory.

## Requirements

- Python 3.8+
- openpyxl ≥ 3.1.0 (see `requirements.txt`)

