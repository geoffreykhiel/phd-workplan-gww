# PhD Workplan — Geoff Whittle-Walls

Python script that generates a fully-formatted Excel workplan for a
**PhD (Direct Entry) · Thesis Incorporating Publications**
at the University of Essex (supervisors: Marie Juanchich & Miro Sirota).

## What it builds

Running the script produces `PhD_Workplan_GWW.xlsx` with five tabs:

| Tab | Contents |
|-----|----------|
| **Workplan** | 150 tasks across 14 phases — planned/actual dates, status, % done, flags |
| **Dashboard** | Summary metrics, phase gate traffic lights, literature review progress |
| **Quality Gates** | Phase-by-phase go/no-go criteria linked to supervisory meetings |
| **Deliverables** | Full outputs register (papers, chapters, thesis milestones) |
| **Literature Review** | 11-column citation log (LR01–LR39) pre-populated with 23 seed papers + keyword search log (Section 2, 20 entries) connected to the Dashboard reading-progress tracker |

## Quick start

```bash
pip install -r requirements.txt
python build_workplan.py
```

The file `PhD_Workplan_GWW.xlsx` is written to the current directory.

## Requirements

- Python 3.8+
- openpyxl ≥ 3.1.0 (see `requirements.txt`)

## Literature Review tab

The **Literature Review** tab follows the 11-column structure agreed with supervisors:

`Ref # · Author(s) · Year · Title · Journal / Source · DOI / URL · Keywords used to find · Stream · Key argument / finding · Relevance to thesis · Status`

**Status dropdown:** `To read / Reading / Read / Cited`
Conditional formatting: green = Read or Cited · amber = Reading

**Section 2 — Keyword Search Log** documents every database search string
(stream, # results, # screened, # retained, forward citation seeds).
This log becomes the literature search methodology section of the thesis.

The Dashboard **Literature Review Progress** panel reads from column K of this tab
to track reading progress in real time.

## Research streams

| Stream | Seed papers |
|--------|-------------|
| Uncertainty communication | LR01–LR03 |
| Risk perception | LR04–LR06, LR13 |
| Truth default theory | LR07–LR09 |
| Vaccine communication | LR10–LR12 |
| Persuasion knowledge | LR14 |
| Dual process theory | LR15–LR16, LR22 |
| Numeracy | LR17, LR23 |
| Framing effects | LR18–LR19 |
| Coherence theory | LR20 |
| Communicative pragmatics | LR21 |
