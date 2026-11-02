#!/usr/bin/env python3
"""
PhD Workplan Tracker — build_workplan.py
=========================================
Generates PhD_Workplan_GWW.xlsx: a multi-tab Excel workplan for a
PhD (Direct Entry) thesis incorporating publications at the University
of Essex (supervisors: Marie Juanchich & Miro Sirota).

Tabs produced
-------------
  Workplan          — 150 tasks across 14 phases with status, dates, flags
  Dashboard         — summary metrics, gate traffic lights, lit-review progress
  Quality Gates     — phase-by-phase go/no-go criteria
  Deliverables      — full outputs register
  Literature Review — 11-column citation log (LR01–LR39) + keyword search log

Usage
-----
  pip install openpyxl
  python build_workplan.py            # writes PhD_Workplan_GWW.xlsx here

Author : Geoff Whittle-Walls
Date   : October 2026
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from datetime import date

OUTPUT = "PhD_Workplan_GWW.xlsx"  # output file is written to the current directory

TITLE_BG = "0F172A"; SECT_BG = "1E3A5F"; GATE_BG = "78350F"; GATE_FG = "FEF3C7"
FINL_BG = "14532D"; FINL_FG = "DCFCE7"; COL_BG = "1E293B"; ROW_EVEN = "F8FAFC"
ROW_ODD = "FFFFFF"; EDIT_BG = "EFF6FF"; TOTL_BG = "1E293B"; DASH_BG = "0F4C81"
LR_BG = "1A3A2A"; LR_SECT = "2D6A4F"; KW_BG = "1B3A5C"

def fill(c): return PatternFill("solid", fgColor=c)
def thin_border():
    s = Side(style="thin", color="CBD5E1")
    return Border(left=s, right=s, top=s, bottom=s)
def cell_font(bold=False, size=10, color="1E293B"):
    return Font(name="Calibri", bold=bold, size=size, color=color)

COLS = [
    ("ID",             6,  False), ("Task",           52, False), ("Area",           16, False),
    ("Planned start",  14, False), ("Planned due",     14, False), ("Actual start",   14, True),
    ("Actual finish",  14, True),  ("Status",          14, True),  ("% done",          9, False),
    ("Days late",      10, False), ("Flag",            22, False), ("Est. hrs",         9, False),
    ("Deliverable",    20, False), ("Notes",           36, True),
]
NC = len(COLS)

SECTIONS = [
    {
        "label": "A.  PRE-REGISTRATION   ·   19 Oct 2026 to 17 Jan 2027",
        "tasks": [
            ("A01","Accept University of Essex offer letter and confirm place",
             "Admin",       date(2026,10,19),date(2026,10,28),  None,""),
            ("A02","Pay tuition deposit and confirm enrolment",
             "Admin",       date(2026,10,19),date(2026,12,14),  None,""),
            ("A03","Complete pre-arrival administrative requirements (IT, student ID, library access)",
             "Admin",       date(2026,11, 1),date(2027, 1,15),  None,""),
            ("A04","Set up citation management system (Zotero) and folder structure for all streams",
             "Research",    date(2026,10,19),date(2026,10,28),   2,  ""),
            ("A05","Conduct systematic literature search – Streams 1–3 (uncertainty communication, risk perception, truth default theory)",
             "Research",    date(2026,10,19),date(2026,11,21),  20,  "Lit review"),
            ("A06","Conduct systematic literature search – Streams 4–6 (vaccine communication & framing, persuasion knowledge, dual process theory)",
             "Research",    date(2026,11, 1),date(2026,12, 5),  20,  "Lit review"),
            ("A07","Conduct systematic literature search – Streams 7–10 (numeracy & risk literacy, framing effects, coherence theory, communicative pragmatics)",
             "Research",    date(2026,11,15),date(2026,12,19),  20,  "Lit review"),
            ("A08","Write synthesis notes across all literature streams – identify gaps and research questions",
             "Research",    date(2026,12, 1),date(2027, 1,10),  15,  "Lit review"),
            ("A09","CTM position paper – complete Steps 2–5 revision",
             "Writing",     date(2026,10,19),date(2026,12,31),  20,  "CTM paper"),
            ("A10","CTM position paper – submit to target journal",
             "Writing",     date(2027, 1, 1),date(2027, 1,15),  None,"CTM paper"),
            ("A11","Draft conceptual framework / theoretical model for thesis",
             "Research",    date(2026,11,15),date(2027, 1,10),  10,  ""),
            ("A12","Identify target journals for Paper 1 and Paper 2 (scope, impact factor, fit)",
             "Research",    date(2026,11, 1),date(2026,12, 1),  None,""),
            ("A13","Draft pilot study research design document (RQ, design, sampling strategy)",
             "Research",    date(2026,11,15),date(2027, 1,10),   8,  "Pilot design"),
            ("A14","Identify and select validated measures for pilot study",
             "Research",    date(2026,11,15),date(2026,12,31),   6,  "Pilot design"),
            ("A15","Conduct power analysis for pilot study (G*Power or equivalent)",
             "Research",    date(2026,12, 1),date(2026,12,31),   3,  "Pilot design"),
            ("A16","Draft pilot study stimulus materials and questionnaire",
             "Research",    date(2026,11,15),date(2027, 1,10),  10,  "Pilot design"),
            ("A17","Draft participant information sheet (PIS) for pilot study",
             "Ethics",      date(2026,12, 1),date(2027, 1,10),   3,  "Ethics app"),
            ("A18","Draft consent form for pilot study",
             "Ethics",      date(2026,12, 1),date(2027, 1,10),   2,  "Ethics app"),
            ("A19","Confirm ethics pathway with Marie – exempt review or full UESS application",
             "Ethics",      date(2026,11,15),date(2026,12,15),  None,"Ethics pathway"),
            ("A20","Pre-registration consultation with Marie – research design, ethics, timeline",
             "Supervision", date(2026,11, 1),date(2026,11,30),  None,""),
            ("A21","Pre-registration consultation with Miro – statistical approach and measures",
             "Supervision", date(2026,12, 1),date(2026,12,31),  None,""),
            ("A22","Contact Essex Doctoral School (Laura) when back from leave – obtain written guidance on what standard PhD regulations permit for including co-authored published work in the thesis; clarify any declaration requirements",
             "Admin",       date(2026,11,15),date(2027, 1,15),  None,"Regulations"),
            ("A23","Clarify with Marie whether the CTM position paper counts toward the thesis papers required by Essex under the direct-entry route; record the agreed position in writing",
             "Supervision", date(2026,11,15),date(2026,12,31),  None,"Thesis"),
            ("A24","Agree authorship order and contribution framework with Marie and Miro for all co-authored outputs to be included in the thesis – record in writing before data collection begins",
             "Supervision", date(2026,11,15),date(2027, 1,15),  None,"Regulations"),
            ("A25","Confirm University of Essex Qualtrics license and account access via Essex IT – required before B08",
             "Admin",       date(2026,11, 1),date(2026,12,15),  None,""),
        ],
        "gate": ("GATE 1","18 Jan 2027",
                 "Official registration at University of Essex  ·  Research activity and ethics applications may begin",
                 False),
    },
    {
        "label": "B.  REGISTRATION + FIRST MONTH   ·   18 Jan 2027 to 28 Feb 2027",
        "tasks": [
            ("B01","Complete official registration at University of Essex",
             "Admin",       date(2027, 1,18),date(2027, 1,18),  None,""),
            ("B02","Set up Doctoral School profile, EDRMS record, and research log",
             "Admin",       date(2027, 1,18),date(2027, 1,25),  None,""),
            ("B03","Complete required Doctoral School training modules (research integrity, data management)",
             "Admin",       date(2027, 1,18),date(2027, 2,28),   6,  ""),
            ("B04","Submit full ethics application to UESS for pilot study",
             "Ethics",      date(2027, 1,18),date(2027, 1,31),   4,  "Ethics app"),
            ("B05","Agree supervision schedule, meeting format, and feedback cadence with Marie and Miro",
             "Supervision", date(2027, 1,18),date(2027, 1,25),  None,""),
            ("B06","Draft Postgraduate Research Plan (if required by Essex Doctoral School)",
             "Admin",       date(2027, 1,18),date(2027, 2,15),   4,  ""),
            ("B07","Finalise pilot study design and questionnaire with Marie's feedback",
             "Research",    date(2027, 1,18),date(2027, 2,15),   5,  "Pilot design"),
            ("B08","Build pilot study data collection infrastructure in Qualtrics",
             "Research",    date(2027, 1,25),date(2027, 2,28),   8,  "Pilot design"),
            ("B09","Draft pilot study debrief materials",
             "Research",    date(2027, 2, 1),date(2027, 2,28),   2,  "Pilot design"),
            ("B10","Draft OSF pre-registration document for pilot study – have fully written and ready to submit immediately on ethics approval; do not register on OSF until C01 is confirmed",
             "Research",    date(2027, 2, 1),date(2027, 3, 1),   3,  "Pilot design"),
            ("B11","Begin Chapter 1 – theoretical framework first draft",
             "Writing",     date(2027, 1,18),date(2027, 2,28),  16,  "Chapter 1"),
            ("B12","January supervision meeting",
             "Supervision", date(2027, 1,18),date(2027, 1,31),  None,""),
            ("B13","Write and submit Data Management Plan (DMP) to Essex Doctoral School – required institutional document; submit within first 4–6 weeks of registration",
             "Admin",       date(2027, 1,18),date(2027, 2,28),   4,  "Regulations"),
            ("B14","GDPR compliance check – confirm data processing agreements with Prolific and Qualtrics; complete Data Protection Impact Assessment (DPIA) if required before any data collection begins",
             "Admin",       date(2027, 1,18),date(2027, 2,28),   3,  "Regulations"),
            ("B15","Identify participant recruitment budget across all three studies (~1,950 participants at Prolific rates) – confirm funding source with Marie and whether Essex covers costs",
             "Admin",       date(2027, 1,18),date(2027, 2,15),  None,"Budget"),
            ("B16","Set up Prolific researcher account, verify payment workflow, and pre-load initial funds – must be operational before C13",
             "Research",    date(2027, 1,18),date(2027, 2,28),   2,  ""),
            ("B17","Establish supervision meeting log in Essex EDRMS from day one – record topics, agreed actions, and any concerns after every meeting; this record is audited at APR",
             "Admin",       date(2027, 1,18),date(2027, 1,25),  None,"APR"),
            ("B18","Identify and register for required Essex Doctoral Skills Programme modules – confirm minimum CPD portfolio requirements with Doctoral School and schedule Year 1 modules now",
             "Admin",       date(2027, 1,18),date(2027, 2,28),   4,  ""),
        ],
        "gate": None,
    },
    {
        "label": "C.  ETHICS APPROVAL + PILOT LAUNCH   ·   1 Mar 2027 to 31 Mar 2027",
        "tasks": [
            ("C01","Receive ethics approval from UESS for pilot study",
             "Ethics",      date(2027, 3, 1),date(2027, 3,15),  None,"Ethics app"),
            ("CBUF","Ethics contingency buffer – 4-week window for UESS queries, amendment requests, or processing delays before committing to pilot launch date",
             "Admin",       date(2027, 3,15),date(2027, 4,12),  None,"Ethics pathway"),
            ("C02","Submit pilot study OSF pre-registration (now ethics is confirmed) – use draft completed at B10",
             "Research",    date(2027, 3, 1),date(2027, 3,15),  None,"Pilot design"),
            ("C03","Launch pilot study – begin participant recruitment via Prolific",
             "Research",    date(2027, 3,15),date(2027, 3,15),  None,"Pilot design"),
            ("C04","Monitor pilot study data collection – check response rates and data quality",
             "Research",    date(2027, 3,15),date(2027, 3,31),   3,  ""),
            ("C05","Chapter 1 – incorporate first-round supervisor feedback, second draft",
             "Writing",     date(2027, 3, 1),date(2027, 3,31),  12,  "Chapter 1"),
            ("C06","Draft Study 1 research design document (RQ, design, sampling)",
             "Research",    date(2027, 3, 1),date(2027, 3,31),   8,  "Study 1"),
            ("C07","Identify and select validated measures for Study 1",
             "Research",    date(2027, 3, 1),date(2027, 3,25),   5,  "Study 1"),
            ("C08","Conduct power analysis for Study 1 sample size",
             "Research",    date(2027, 3,15),date(2027, 3,31),   3,  "Study 1"),
            ("C09","Draft Study 1 participant information sheet (PIS)",
             "Ethics",      date(2027, 3,15),date(2027, 3,31),   3,  "Ethics app"),
            ("C10","Draft Study 1 consent form",
             "Ethics",      date(2027, 3,15),date(2027, 3,31),   2,  "Ethics app"),
            ("C11","February supervision meeting",
             "Supervision", date(2027, 2, 1),date(2027, 2,28),  None,""),
            ("C12","March supervision meeting",
             "Supervision", date(2027, 3, 1),date(2027, 3,31),  None,""),
            ("C13","Pilot pre-launch technical test on Prolific – run 10–20 paid participants through the complete survey flow before full launch; verify redirect links, completion codes, quota-balanced sampling, and attention-check branching",
             "Research",    date(2027, 3,12),date(2027, 3,15),   3,  "Pilot design"),
        ],
        "gate": ("GATE 2","15 Mar 2027",
                 "Ethics approval received from UESS  ·  Primary data collection may begin",
                 False),
    },
    {
        "label": "D.  PILOT STUDY CLOSEOUT + STUDY 1 DESIGN   ·   1 Apr 2027 to 30 Apr 2027",
        "tasks": [
            ("D01","Pilot study data collection complete",
             "Research",    date(2027, 4, 1),date(2027, 4,10),  None,"Pilot design"),
            ("D02","Pilot study data cleaning and quality check (exclusions, attention checks)",
             "Research",    date(2027, 4, 1),date(2027, 4,15),   4,  "Pilot write-up"),
            ("D03","Pilot study analysis – pre-specified and exploratory",
             "Research",    date(2027, 4, 1),date(2027, 4,20),  10,  "Pilot write-up"),
            ("D04","Pilot study findings summary and write-up (internal)",
             "Research",    date(2027, 4,15),date(2027, 4,30),   6,  "Pilot write-up"),
            ("D05","Amend ethics application or submit amendment if pilot findings require changes to Study 1 protocol",
             "Ethics",      date(2027, 4,15),date(2027, 4,30),   3,  "Ethics app"),
            ("D06","Finalise Study 1 design incorporating pilot findings",
             "Research",    date(2027, 4, 1),date(2027, 4,30),   8,  "Study 1"),
            ("D07","Build Study 1 data collection infrastructure in Qualtrics",
             "Research",    date(2027, 4,15),date(2027, 4,30),   8,  "Study 1"),
            ("D08","Submit Study 1 ethics application (or amendment) to UESS",
             "Ethics",      date(2027, 4,15),date(2027, 4,30),   4,  "Ethics app"),
            ("D09","Write detailed Statistical Analysis Plan (SAP) for Study 1 – specify primary/secondary outcomes, model, exclusion criteria, multiple comparisons approach, and sensitivity analyses; include in OSF pre-registration",
             "Research",    date(2027, 4,15),date(2027, 4,30),   5,  "Study 1"),
            ("D10","Pre-register Study 1 on OSF (after ethics approval or amendment confirmed)",
             "Research",    date(2027, 4,25),date(2027, 5,15),   2,  "Study 1"),
            ("D11","Chapter 1 – incorporate second-round supervisor comments",
             "Writing",     date(2027, 4, 1),date(2027, 4,30),   8,  "Chapter 1"),
            ("D12","April supervision meeting",
             "Supervision", date(2027, 4, 1),date(2027, 4,30),  None,""),
            ("D13","Study 1 pre-launch technical test on Prolific – 10–20 paid participants to verify Qualtrics–Prolific integration before full E02 launch",
             "Research",    date(2027, 5, 5),date(2027, 5,10),   2,  "Study 1"),
        ],
        "gate": None,
    },
    {
        "label": "E.  STUDY 1 DATA COLLECTION   ·   1 May 2027 to 30 Jun 2027",
        "tasks": [
            ("E01","Receive Study 1 ethics clearance (or amendment clearance)",
             "Ethics",      date(2027, 5, 1),date(2027, 5,10),  None,"Ethics app"),
            ("E02","Launch Study 1 – begin participant recruitment via Prolific",
             "Research",    date(2027, 5,10),date(2027, 5,10),  None,"Study 1"),
            ("E03","Monitor Study 1 data collection – response rates, data quality, and dropout",
             "Research",    date(2027, 5,10),date(2027, 6,25),   5,  ""),
            ("E04","Study 1 data collection complete",
             "Research",    date(2027, 6,25),date(2027, 6,30),  None,"Study 1"),
            ("E05","Complete Chapter 1 – final draft ready for supervisor sign-off",
             "Writing",     date(2027, 5, 1),date(2027, 5,31),  10,  "Chapter 1"),
            ("E06","Begin Chapter 2 – empirical overview, methods, and analytical strategy",
             "Writing",     date(2027, 6, 1),date(2027, 6,30),   8,  "Chapter 2"),
            ("E07","Submit abstract to at least one peer conference (APA, SPSP, BPS, or equivalent)",
             "Writing",     date(2027, 5, 1),date(2027, 5,31),   3,  "Conference"),
            ("E08","Begin drafting Study 2 design (based on Study 1 design; can start early)",
             "Research",    date(2027, 6, 1),date(2027, 6,30),   5,  "Study 2"),
            ("E09","May supervision meeting",
             "Supervision", date(2027, 5, 1),date(2027, 5,31),  None,""),
            ("E10","June supervision meeting",
             "Supervision", date(2027, 6, 1),date(2027, 6,30),  None,""),
            ("E11","Conference attendance and presentation – if E07 abstract is accepted, book registration, accommodation, and travel; prepare and deliver presentation",
             "Writing",     date(2027, 8, 1),date(2027,10,31),   8,  "Conference"),
        ],
        "gate": None,
    },
    {
        "label": "F.  STUDY 1 ANALYSIS + STUDY 2 DATA COLLECTION   ·   1 Jul 2027 to 31 Aug 2027",
        "tasks": [
            ("F01","Study 1 data cleaning and quality check (exclusions, attention checks, outliers)",
             "Research",    date(2027, 7, 1),date(2027, 7,10),   5,  "Study 1"),
            ("F02","Study 1 pre-specified confirmatory analysis",
             "Research",    date(2027, 7,10),date(2027, 7,20),  10,  "Study 1"),
            ("F03","Study 1 exploratory analyses and sensitivity checks",
             "Research",    date(2027, 7,20),date(2027, 7,31),   8,  "Study 1"),
            ("F04","Finalise Study 2 design incorporating Study 1 findings",
             "Research",    date(2027, 7, 1),date(2027, 7,15),   8,  "Study 2"),
            ("F05","Draft Study 2 participant information sheet (PIS)",
             "Ethics",      date(2027, 7, 1),date(2027, 7,15),   3,  "Ethics app"),
            ("F06","Draft Study 2 consent form",
             "Ethics",      date(2027, 7, 1),date(2027, 7,15),   2,  "Ethics app"),
            ("F07","Submit Study 2 ethics application to UESS (if new or substantially amended protocol)",
             "Ethics",      date(2027, 7,10),date(2027, 7,20),   4,  "Ethics app"),
            ("F08","Build Study 2 data collection infrastructure in Qualtrics",
             "Research",    date(2027, 7,15),date(2027, 7,31),   6,  "Study 2"),
            ("F09","Write detailed Statistical Analysis Plan (SAP) for Study 2 – specify models, exclusions, multiple comparisons approach; include in OSF pre-registration",
             "Research",    date(2027, 7,15),date(2027, 7,25),   4,  "Study 2"),
            ("F10","Pre-register Study 2 on OSF (after ethics approval)",
             "Research",    date(2027, 7,20),date(2027, 8,10),   2,  "Study 2"),
            ("F11","Study 2 pre-launch technical test on Prolific – 10–20 paid participants to verify Qualtrics–Prolific integration before full launch",
             "Research",    date(2027, 7,28),date(2027, 8, 1),   2,  "Study 2"),
            ("F12","Receive Study 2 ethics clearance and launch Study 2 via Prolific",
             "Research",    date(2027, 7,25),date(2027, 8, 1),  None,"Study 2"),
            ("F13","Monitor Study 2 data collection – response rates and data quality",
             "Research",    date(2027, 8, 1),date(2027, 8,25),   4,  ""),
            ("F14","Study 2 data collection complete",
             "Research",    date(2027, 8,25),date(2027, 8,31),  None,"Study 2"),
            ("F15","Paper 1 first draft – Study 1 findings (introduction, methods, results, discussion)",
             "Writing",     date(2027, 7, 1),date(2027, 8,15),  25,  "Paper 1"),
            ("F16","Chapter 2 – complete draft (empirical overview, methods, analysis approach)",
             "Writing",     date(2027, 7, 1),date(2027, 8,31),  15,  "Chapter 2"),
            ("F17","July supervision meeting",
             "Supervision", date(2027, 7, 1),date(2027, 7,31),  None,""),
            ("F18","August supervision meeting",
             "Supervision", date(2027, 8, 1),date(2027, 8,31),  None,""),
            ("F19","Prepare Study 1 open data package – cleaned dataset, codebook, analysis scripts, README; upload to OSF while analysis is fresh",
             "Research",    date(2027, 8, 1),date(2027, 8,31),   5,  "Study 1"),
        ],
        "gate": ("GATE 3","30 Sep 2027",
                 "Annual Progress Review passed  ·  Continuation to examination confirmed",
                 False),
    },
    {
        "label": "G.  ANNUAL REVIEW + STUDY 2 ANALYSIS   ·   1 Sep 2027 to 31 Oct 2027",
        "tasks": [
            ("G01","Study 2 data cleaning and quality check",
             "Research",    date(2027, 9, 1),date(2027, 9,10),   5,  "Study 2"),
            ("G02","Study 2 pre-specified confirmatory analysis",
             "Research",    date(2027, 9,10),date(2027, 9,20),  10,  "Study 2"),
            ("G03","Study 2 exploratory analyses and sensitivity checks",
             "Research",    date(2027, 9,20),date(2027, 9,30),   8,  "Study 2"),
            ("G04","Paper 1 – incorporate Marie and Miro's feedback (first revision)",
             "Writing",     date(2027, 9, 1),date(2027, 9,20),  10,  "Paper 1"),
            ("G05","Write journal cover letter for Paper 1 – position contribution for editors; make scope and theoretical claim explicit",
             "Writing",     date(2027, 9,15),date(2027,10,10),   3,  "Paper 1"),
            ("G06","Submit Paper 1 to target journal (with cover letter from G05)",
             "Writing",     date(2027, 9,20),date(2027,10,15),  None,"Paper 1"),
            ("G07","Draft APR written report – progress to date, revised timeline, any obstacles",
             "Admin",       date(2027, 9, 1),date(2027, 9,25),   8,  "APR"),
            ("G08","Annual Progress Review – submit written report to supervisors and panel",
             "Admin",       date(2027, 9,25),date(2027, 9,30),  None,"APR"),
            ("G09","Annual Progress Review – panel meeting with Marie, Miro and Graduate Director",
             "Supervision", date(2027,10, 1),date(2027,10,14),  None,"APR"),
            ("G10","Implement APR panel recommendations and submit written response to any conditions or required actions – required if panel issues formal conditions",
             "Admin",       date(2027,10,14),date(2027,10,31),   4,  "APR"),
            ("G11","Paper 2 first draft – Study 2 findings (introduction, methods, results, discussion)",
             "Writing",     date(2027, 9, 1),date(2027,10,31),  25,  "Paper 2"),
            ("G12","Chapter 3 – first draft (discussion and theoretical integration across studies)",
             "Writing",     date(2027, 9, 1),date(2027,10,31),  20,  "Chapter 3"),
            ("G13","Prepare Study 2 open data package – cleaned dataset, codebook, analysis scripts, README; upload to OSF while analysis is fresh",
             "Research",    date(2027, 9,20),date(2027,10,31),   5,  "Study 2"),
            ("G14","September supervision meeting",
             "Supervision", date(2027, 9, 1),date(2027, 9,30),  None,""),
            ("G15","October supervision meeting",
             "Supervision", date(2027,10, 1),date(2027,10,31),  None,""),
        ],
        "gate": None,
    },
    {
        "label": "H.  FINAL WRITE-UP   ·   1 Nov 2027 to 31 Dec 2027",
        "tasks": [
            ("H01","Paper 2 – incorporate Marie and Miro's feedback (first revision)",
             "Writing",     date(2027,11, 1),date(2027,11,20),  10,  "Paper 2"),
            ("H02","Write journal cover letter for Paper 2 – position contribution and scope for editors",
             "Writing",     date(2027,11,15),date(2027,11,25),   3,  "Paper 2"),
            ("H03","Submit Paper 2 to target journal (with cover letter from H02)",
             "Writing",     date(2027,11,20),date(2027,11,30),  None,"Paper 2"),
            ("H04","Paper 1 – respond to journal peer review comments and resubmit; include formal point-by-point response letter to editor",
             "Writing",     date(2027,11, 1),date(2027,12,15),  15,  "Paper 1"),
            ("H05","Chapter 3 – incorporate supervisor feedback, second draft",
             "Writing",     date(2027,11, 1),date(2027,11,30),  12,  "Chapter 3"),
            ("H06","Chapter 4 – first draft (conclusions, contribution to knowledge, future directions)",
             "Writing",     date(2027,11, 1),date(2027,11,30),  15,  "Chapter 4"),
            ("H07","Write thesis abstract – 800–1,000 words covering background, methods, key findings across studies, and overall contribution to knowledge",
             "Writing",     date(2027,11,15),date(2027,12,15),   5,  "Thesis"),
            ("H08","Draft candidate's contribution statement for each co-authored paper to be included in thesis – obtain signed copies from Marie and Miro; required under Essex standard PhD regulations for including published co-authored work",
             "Admin",       date(2027,11, 1),date(2027,12,15),   3,  "Regulations"),
            ("H09","Send full thesis draft (all chapters, abstract, and declarations) to Marie and Miro",
             "Writing",     date(2027,11,30),date(2027,12, 5),  None,"Thesis"),
            ("H10","Incorporate supervisor feedback into full thesis draft",
             "Writing",     date(2027,12, 5),date(2027,12,20),  25,  "Thesis"),
            ("H11","Final bibliography audit in Zotero – check for duplicates, formatting consistency, and completeness of all in-text citations before submission",
             "Writing",     date(2027,12,15),date(2027,12,28),   4,  "Thesis"),
            ("H12","Run Turnitin originality check on complete thesis draft – required before submission",
             "Admin",       date(2027,12,20),date(2027,12,28),   1,  "Thesis"),
            ("H13","Check and apply University of Essex thesis formatting requirements",
             "Admin",       date(2027,12,15),date(2027,12,28),   5,  "Thesis"),
            ("H14","Complete final integrated thesis draft",
             "Writing",     date(2027,12,20),date(2027,12,31),  10,  "Thesis"),
            ("H15","Order hard-bound thesis copies from professional print/bind service – allow 10–14 working days minimum",
             "Admin",       date(2027,12,20),date(2027,12,28),  None,"Thesis"),
            ("H16","Agree internal and external examiner nominations with Marie and Miro – submit to Doctoral School for approval (required before submission)",
             "Admin",       date(2027,12, 1),date(2027,12,15),  None,"Viva"),
            ("H17","November supervision meeting",
             "Supervision", date(2027,11, 1),date(2027,11,30),  None,""),
            ("H18","December supervision meeting",
             "Supervision", date(2027,12, 1),date(2027,12,31),  None,""),
        ],
        "gate": None,
    },
    {
        "label": "I.  SUBMISSION   ·   2 Jan 2028 to 18 Jan 2028",
        "tasks": [
            ("I01","Collect hard-bound copies from binder – confirm with Essex Doctoral School whether physical copies are required at submission or post-viva",
             "Admin",       date(2028, 1, 2),date(2028, 1,10),  None,"Thesis"),
            ("I02","Final thesis polish – read-through, consistency check, and proofreading",
             "Writing",     date(2028, 1, 2),date(2028, 1,12),  10,  "Thesis"),
            ("I03","Verify thesis meets Essex word count, formatting, declaration, and abstract requirements",
             "Admin",       date(2028, 1, 2),date(2028, 1,14),   3,  "Thesis"),
            ("I04","Pre-submission sign-off from Marie and Miro",
             "Supervision", date(2028, 1,10),date(2028, 1,15),  None,"Thesis"),
            ("I05","Upload and submit thesis electronically through Essex submission portal",
             "Admin",       date(2028, 1,15),date(2028, 1,18),  None,"Thesis"),
            ("I06","Confirm receipt of submission and log completion",
             "Admin",       date(2028, 1,18),date(2028, 1,18),  None,""),
        ],
        "gate": ("GATE 4","18 Jan 2028",
                 "Thesis submitted for examination  ·  Viva preparation begins",
                 False),
    },
    {
        "label": "J.  VIVA PREPARATION + EXAMINATION   ·   Jan 2028 onwards",
        "tasks": [
            ("J01","Examiners receive thesis and begin reading period – no action required; remain available for any examiner queries via Doctoral School",
             "Admin",       date(2028, 1,18),date(2028, 2,28),  None,"Viva"),
            ("J02","Viva date confirmed by Doctoral School – check availability and arrange travel if required",
             "Admin",       date(2028, 1,18),date(2028, 2,28),  None,"Viva"),
            ("J03","Viva preparation – structured review: re-read papers and all thesis chapters; map each examiner's published work to your theoretical framework; draft responses to likely lines of questioning (SDT, intention inference, generalisability, limitations)",
             "Research",    date(2028, 1,18),date(2028, 3,14),  15,  "Viva"),
            ("J04","Mock viva with Marie and Miro – full rehearsal under exam conditions; agree key messages and framing of thesis contribution",
             "Supervision", date(2028, 3, 1),date(2028, 3,14),   4,  "Viva"),
            ("J05","Viva voce examination – oral defence of thesis before internal and external examiner",
             "Supervision", date(2028, 3,14),date(2028, 3,28),  None,"Viva"),
            ("J06","Receive examiner outcome and written corrections report – review with Marie; agree scope, priority, and timeline for revisions",
             "Supervision", date(2028, 3,28),date(2028, 4,14),  None,"Viva"),
            ("J07","Complete and resubmit corrections (minor revisions) to Doctoral School",
             "Writing",     date(2028, 4, 1),date(2028, 5,31),  20,  "Thesis"),
            ("J08","Update any papers in thesis to final accepted/published versions with confirmed DOIs – if decisions arrive before or shortly after viva",
             "Writing",     date(2028, 4, 1),date(2028, 5,31),   4,  "Thesis"),
            ("J09","Corrections accepted – formal notification of award from University of Essex",
             "Admin",       date(2028, 5,31),date(2028, 6,30),  None,""),
            ("J10","Deposit corrected thesis in Essex Research Online (ERO) and British Library EThOS – institutional requirement following award confirmation",
             "Admin",       date(2028, 6, 1),date(2028, 6,30),   2,  ""),
            ("J11","Update ORCID, Google Scholar, ResearchGate, and Essex academic profile with PhD award, thesis, and final paper citations",
             "Admin",       date(2028, 6, 1),date(2028, 6,30),   2,  ""),
        ],
        "gate": ("FINAL","Jun 2028",
                 "Corrections accepted  ·  PhD (thesis incorporating publications) awarded  ·  University of Essex",
                 True),
    },
]

# ── Build workbook ─────────────────────────────────────────────────────────────
wb = Workbook()
ws_wp = wb.active
ws_wp.title = "Workplan"
last_col = get_column_letter(NC)

for i, (_, w, _) in enumerate(COLS, 1):
    ws_wp.column_dimensions[get_column_letter(i)].width = w

ws_wp.row_dimensions[1].height = 26
ws_wp.merge_cells(f"A1:{last_col}1")
c = ws_wp["A1"]
c.value = "PhD (Direct Entry)  ·  Thesis Incorporating Publications  ·  Workplan Tracker"
c.fill = fill(TITLE_BG)
c.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

ws_wp.row_dimensions[2].height = 16
ws_wp.merge_cells(f"A2:{last_col}2")
c = ws_wp["A2"]
c.value = ("Geoff Whittle-Walls  ·  University of Essex  ·  Psychology  ·  "
           "Supervisors: Marie Juanchich & Miro Sirota  ·  18 Jan 2027 to 18 Jan 2028  ·  "
           "Standard PhD regulations apply  ·  Shaded cells are the only ones you edit")
c.fill = fill(TITLE_BG)
c.font = Font(name="Calibri", size=10, color="94A3B8")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

ws_wp.row_dimensions[3].height = 22
for i, (hdr, _, _) in enumerate(COLS, 1):
    c = ws_wp.cell(row=3, column=i, value=hdr)
    c.fill = fill(COL_BG)
    c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = thin_border()

ws_wp.freeze_panes = "A4"

DATE_FMT = "D MMM YYYY"
row = 4; data_rows = []; even_row = True

for sec in SECTIONS:
    ws_wp.row_dimensions[row].height = 17
    ws_wp.merge_cells(f"A{row}:{last_col}{row}")
    c = ws_wp.cell(row=row, column=1, value=sec["label"])
    c.fill = fill(SECT_BG)
    c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row += 1

    for t in sec["tasks"]:
        tid, task, area, pstart, pdue, hrs, deliv = t
        r = row; data_rows.append(r); row_bg = ROW_EVEN if even_row else ROW_ODD
        even_row = not even_row
        ws_wp.row_dimensions[r].height = 16
        static = {1:tid, 2:task, 3:area, 4:pstart, 5:pdue, 8:"Not started",
                  12:hrs if hrs else None, 13:deliv if deliv else None}
        for ci in range(1, NC+1):
            cell = ws_wp.cell(row=r, column=ci)
            _, _, is_edit = COLS[ci-1]; val = static.get(ci)
            if val is not None: cell.value = val
            if ci in (4, 5) and isinstance(val, date): cell.number_format = DATE_FMT
            cell.fill = fill(EDIT_BG if is_edit else row_bg)
            cell.font = cell_font(size=10)
            cell.alignment = Alignment(vertical="center",
                horizontal="left" if ci in (2, 14) else "center")
            cell.border = thin_border()
        ws_wp.cell(row=r, column=6).number_format = DATE_FMT
        ws_wp.cell(row=r, column=7).number_format = DATE_FMT

        c9 = ws_wp.cell(row=r, column=9)
        c9.value = f'=IF(H{r}="Complete",1,IF(H{r}="In progress",0.5,0))'
        c9.number_format = "0%"; c9.fill = fill(row_bg); c9.font = cell_font(size=10)
        c9.alignment = Alignment(horizontal="center", vertical="center"); c9.border = thin_border()

        c10 = ws_wp.cell(row=r, column=10)
        c10.value = (f'=IF(AND(H{r}<>"Complete",H{r}<>"",E{r}<>"",E{r}<TODAY()),'
                     f'TODAY()-E{r},0)')
        c10.fill = fill(row_bg); c10.font = cell_font(size=10)
        c10.alignment = Alignment(horizontal="center", vertical="center"); c10.border = thin_border()

        c11 = ws_wp.cell(row=r, column=11)
        c11.value = (
            f'=IF(H{r}="Blocked","BLOCKED",'
            f'IF(AND(H{r}<>"Complete",H{r}<>"",E{r}<>"",E{r}<TODAY()),"OVERDUE",'
            f'IF(AND(H{r}<>"Complete",H{r}<>"",E{r}<>"",E{r}-TODAY()<=5,E{r}>=TODAY()),"DUE SOON",'
            f'IF(AND(H{r}="Not started",D{r}<>"",D{r}<TODAY()),"SHOULD HAVE STARTED",""))))'
        )
        c11.fill = fill(row_bg); c11.font = cell_font(bold=True, size=10)
        c11.alignment = Alignment(horizontal="center", vertical="center"); c11.border = thin_border()
        row += 1

    if sec["gate"]:
        glabel, gdate, gdesc, is_final = sec["gate"]
        ws_wp.row_dimensions[row].height = 18
        ws_wp.merge_cells(f"A{row}:{last_col}{row}")
        gbg = FINL_BG if is_final else GATE_BG; gfg = FINL_FG if is_final else GATE_FG
        c = ws_wp.cell(row=row, column=1,
                       value=f"   {glabel}  ·  {gdate}  ·  {gdesc}")
        c.fill = fill(gbg); c.font = Font(name="Calibri", bold=True, size=10, color=gfg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        row += 1

ws_wp.row_dimensions[row].height = 17
for ci in range(1, NC+1):
    cell = ws_wp.cell(row=row, column=ci)
    cell.fill = fill(TOTL_BG); cell.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    cell.border = thin_border()
    cell.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")

ws_wp.cell(row=row, column=1).value = "TOTAL"
ws_wp.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)

first_dr, last_dr = data_rows[0], data_rows[-1]
pct_total = (
    f'=(COUNTIF(H{first_dr}:H{last_dr},"Complete")'
    f'+COUNTIF(H{first_dr}:H{last_dr},"In progress")*0.5)'
    f'/MAX(1,COUNTIF(H{first_dr}:H{last_dr},"Not started")'
    f'+COUNTIF(H{first_dr}:H{last_dr},"In progress")'
    f'+COUNTIF(H{first_dr}:H{last_dr},"Complete")'
    f'+COUNTIF(H{first_dr}:H{last_dr},"Blocked"))'
)
ws_wp.cell(row=row, column=9).value = pct_total
ws_wp.cell(row=row, column=9).number_format = "0%"
ws_wp.cell(row=row, column=9).fill = fill(TOTL_BG)
ws_wp.cell(row=row, column=9).font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
ws_wp.cell(row=row, column=9).border = thin_border()
ws_wp.cell(row=row, column=9).alignment = Alignment(horizontal="center", vertical="center")
ws_wp.cell(row=row, column=12).value = f'=SUM(L{first_dr}:L{last_dr})'
ws_wp.cell(row=row, column=12).fill = fill(TOTL_BG)
ws_wp.cell(row=row, column=12).font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
ws_wp.cell(row=row, column=12).border = thin_border()
ws_wp.cell(row=row, column=12).alignment = Alignment(horizontal="center", vertical="center")

dv = DataValidation(type="list",
    formula1='"Not started,In progress,Complete,Blocked"',
    allow_blank=True, showDropDown=False,
    showInputMessage=True, promptTitle="Status",
    prompt="Choose: Not started / In progress / Complete / Blocked")
dv.sqref = f"H{first_dr}:H{last_dr}"
ws_wp.add_data_validation(dv)

flag_range = f"K{first_dr}:K{last_dr}"
for formula, bg, fg in [
    (f'$K{first_dr}="OVERDUE"',             "FCA5A5","991B1B"),
    (f'$K{first_dr}="DUE SOON"',            "FCD34D","92400E"),
    (f'$K{first_dr}="BLOCKED"',             "FCA5A5","991B1B"),
    (f'$K{first_dr}="SHOULD HAVE STARTED"', "FED7AA","92400E"),
]:
    ws_wp.conditional_formatting.add(flag_range,
        FormulaRule(formula=[formula],
            fill=PatternFill("solid", fgColor=bg),
            font=Font(bold=True, color=fg), stopIfTrue=False))

print(f"✓ Workplan: {len(data_rows)} tasks, rows {first_dr}–{last_dr}")

# ═══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
ws_db = wb.create_sheet("Dashboard")
ws_db.sheet_view.showGridLines = False
ws_db.column_dimensions["A"].width = 38
ws_db.column_dimensions["B"].width = 18
ws_db.column_dimensions["C"].width = 16
ws_db.column_dimensions["D"].width = 14
ws_db.column_dimensions["E"].width = 14
ws_db.column_dimensions["F"].width = 14
ws_db.column_dimensions["G"].width = 14
ws_db.column_dimensions["H"].width = 32

def db_banner(ws, r, text, bg=TITLE_BG, fg="FFFFFF", size=13, ht=26):
    ws.row_dimensions[r].height = ht
    ws.merge_cells(f"A{r}:H{r}")
    c = ws.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = Font(name="Calibri", bold=True, size=size, color=fg)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def db_sect(ws, r, text, height=17):
    ws.row_dimensions[r].height = height
    ws.merge_cells(f"A{r}:H{r}")
    c = ws.cell(row=r, column=1, value=text)
    c.fill = fill(DASH_BG); c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def db_hdr(ws, r, hdrs, height=18):
    ws.row_dimensions[r].height = height
    for ci, txt in enumerate(hdrs, 1):
        c = ws.cell(row=r, column=ci, value=txt)
        c.fill = fill(DASH_BG); c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border()

def db_data(ws, r, vals, fmts=None, bgs=None, bold1=False, height=16):
    ws.row_dimensions[r].height = height
    for ci, val in enumerate(vals, 1):
        c = ws.cell(row=r, column=ci, value=val)
        bg = bgs[ci-1] if bgs else (ROW_EVEN if r % 2 == 0 else ROW_ODD)
        c.fill = fill(bg)
        c.font = Font(name="Calibri", bold=(bold1 and ci == 1), size=10, color="1E293B")
        c.alignment = Alignment(horizontal="left" if ci == 1 else "center", vertical="center")
        c.border = thin_border()
        if fmts and fmts[ci-1]: c.number_format = fmts[ci-1]

db_banner(ws_db, 1, "Where the programme stands")
ws_db.row_dimensions[2].height = 16
ws_db.merge_cells("A2:H2")
c = ws_db.cell(row=2, column=1,
               value="Everything here calculates from the Workplan sheet. Nothing is typed by hand.")
c.fill = fill(TITLE_BG); c.font = Font(name="Calibri", size=10, color="94A3B8")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

db_sect(ws_db, 4, "Programme overview")
metrics = [
    ("Today",                        "=TODAY()",                                    "D MMM YYYY",""),
    ("Registration start",           "18 Jan 2027",                                 "D MMM YYYY",""),
    ("Thesis submission deadline",   "18 Jan 2028",                                 "D MMM YYYY","Fixed. Work backwards from this date."),
    ("Working days left to deadline","=NETWORKDAYS(TODAY(),DATE(2028,1,18))-1",     "0",""),
    ("Working days in programme",    "=NETWORKDAYS(DATE(2027,1,18),DATE(2028,1,18))-1","0",""),
    ("Elapsed, share of period",
     "=MAX(0,(TODAY()-DATE(2027,1,18))/MAX(1,DATE(2028,1,18)-DATE(2027,1,18)))",   "0%",""),
    ("Work complete, share of tasks", pct_total,                                    "0%",""),
    ("Schedule position",
     f'=({pct_total})-MAX(0,(TODAY()-DATE(2027,1,18))/MAX(1,DATE(2028,1,18)-DATE(2027,1,18)))',
     "+0%;-0%;0%","Positive = ahead; negative = behind the calendar"),
]
for i, (label, formula, fmt, note) in enumerate(metrics):
    r2 = 5 + i
    ws_db.row_dimensions[r2].height = 16
    bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    for ci in range(1, 9):
        c = ws_db.cell(row=r2, column=ci)
        c.fill = fill(bg); c.border = thin_border()
        c.alignment = Alignment(horizontal="left" if ci in (1,8) else "center", vertical="center")
        c.font = Font(name="Calibri", size=10)
    ws_db.cell(row=r2, column=1).value = label
    ws_db.cell(row=r2, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2 = ws_db.cell(row=r2, column=2, value=formula)
    c2.number_format = fmt; c2.font = Font(name="Calibri", bold=True, size=10)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    if note:
        c8 = ws_db.cell(row=r2, column=8, value=note)
        c8.font = Font(name="Calibri", size=9, color="64748B", italic=True)

br = 5 + len(metrics) + 2
db_sect(ws_db, br, "By block")
br += 1
db_hdr(ws_db, br, ["Block","Tasks","Complete","In progress","Blocked","Not started","% done","Est. hrs"])

blocks = [
    ("A. Pre-registration",            "A"),
    ("B. Registration + first month",  "B"),
    ("C. Ethics approval + pilot",     "C"),
    ("D. Pilot closeout + Study 1",    "D"),
    ("E. Study 1 data collection",     "E"),
    ("F. Study 1 analysis + Study 2",  "F"),
    ("G. Annual review + Study 2",     "G"),
    ("H. Final write-up",              "H"),
    ("I. Submission",                  "I"),
    ("J. Viva + examination",          "J"),
]
for i, (blabel, bid) in enumerate(blocks):
    br += 1
    n_tasks = (f'=COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Not started")'
               f'+COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"In progress")'
               f'+COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Complete")'
               f'+COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Blocked")')
    n_comp = f'=COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Complete")'
    n_inp  = f'=COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"In progress")'
    n_blk  = f'=COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Blocked")'
    n_ns   = f'=COUNTIFS(Workplan!A:A,"{bid}*",Workplan!H:H,"Not started")'
    n_pct  = f'=IFERROR(({n_comp}+{n_inp}*0.5)/MAX(1,{n_tasks}),0)'
    n_hrs  = f'=SUMIF(Workplan!A:A,"{bid}*",Workplan!L:L)'
    bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    db_data(ws_db, br, [blabel, n_tasks, n_comp, n_inp, n_blk, n_ns, n_pct, n_hrs],
            fmts=[None,"0","0","0","0","0","0%","0"], bgs=[bg]*8, bold1=True)

br += 1
ws_db.row_dimensions[br].height = 17
for ci in range(1, 9):
    c = ws_db.cell(row=br, column=ci)
    c.fill = fill(TOTL_BG); c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    c.border = thin_border(); c.alignment = Alignment(horizontal="center" if ci>1 else "left", vertical="center")
ws_db.cell(row=br, column=1).value = "TOTAL"
ws_db.cell(row=br, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
ttasks = (f'=COUNTIF(Workplan!H:H,"Not started")+COUNTIF(Workplan!H:H,"In progress")'
          f'+COUNTIF(Workplan!H:H,"Complete")+COUNTIF(Workplan!H:H,"Blocked")')
for ci, (val, fmt) in enumerate([
    (ttasks,"0"), (f'=COUNTIF(Workplan!H:H,"Complete")',"0"),
    (f'=COUNTIF(Workplan!H:H,"In progress")',"0"),
    (f'=COUNTIF(Workplan!H:H,"Blocked")',"0"),
    (f'=COUNTIF(Workplan!H:H,"Not started")',"0"),
    (pct_total,"0%"), (f'=SUM(Workplan!L:L)',"0"),
], start=2):
    c = ws_db.cell(row=br, column=ci, value=val)
    c.number_format = fmt; c.fill = fill(TOTL_BG)
    c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    c.border = thin_border(); c.alignment = Alignment(horizontal="center", vertical="center")

ar = br + 2
db_sect(ws_db, ar, "Attention now")
ar += 1
db_hdr(ws_db, ar, ["","Count","","","","","","What it means"])
attn = [
    ("Overdue",
     f'=SUMPRODUCT((Workplan!E{first_dr}:E{last_dr}<TODAY())*(Workplan!H{first_dr}:H{last_dr}<>"Complete")*(Workplan!H{first_dr}:H{last_dr}<>"")*(Workplan!E{first_dr}:E{last_dr}<>""))',
     "Past the planned due date and not yet marked Complete"),
    ("Blocked",
     f'=COUNTIF(Workplan!H{first_dr}:H{last_dr},"Blocked")',
     "Waiting on something external – chase these first"),
    ("Should have started",
     f'=SUMPRODUCT((Workplan!D{first_dr}:D{last_dr}<TODAY())*(Workplan!H{first_dr}:H{last_dr}="Not started")*(Workplan!D{first_dr}:D{last_dr}<>""))',
     "Start date has passed but status is still Not started"),
    ("Due in next five working days",
     f'=SUMPRODUCT((Workplan!E{first_dr}:E{last_dr}>=TODAY())*(Workplan!E{first_dr}:E{last_dr}<=TODAY()+7)*(Workplan!H{first_dr}:H{last_dr}<>"Complete")*(Workplan!E{first_dr}:E{last_dr}<>""))',
     "Imminent – make sure these are actively in progress"),
]
for i, (label, formula, note) in enumerate(attn):
    ar += 1
    bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    ws_db.row_dimensions[ar].height = 16
    for ci in range(1, 9):
        c = ws_db.cell(row=ar, column=ci)
        c.fill = fill(bg); c.border = thin_border()
        c.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")
        c.font = Font(name="Calibri", size=10)
    ws_db.cell(row=ar, column=1).value = label
    ws_db.cell(row=ar, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2 = ws_db.cell(row=ar, column=2, value=formula)
    c2.number_format = "0"; c2.font = Font(name="Calibri", bold=True, size=10)
    c8 = ws_db.cell(row=ar, column=8, value=note)
    c8.font = Font(name="Calibri", size=9, color="64748B", italic=True)

# ── Literature Review Progress ─────────────────────────────────────────────────
ar += 2
db_sect(ws_db, ar, "Literature Review Progress")
ar += 1
db_hdr(ws_db, ar, ["", "Count", "", "", "", "", "", "What it means"])
lit_progress = [
    ("Total papers logged",
     "=COUNTA('Literature Review'!B5:B43)",
     "Papers entered in the Citation Log (LR01–LR39)"),
    ("Read",
     '=COUNTIF(\'Literature Review\'!K5:K43,"Read")',
     "Papers with 'Read' in the Status column"),
    ("Cited",
     '=COUNTIF(\'Literature Review\'!K5:K43,"Cited")',
     "Papers marked Cited — incorporated into thesis writing"),
    ("Reading",
     '=COUNTIF(\'Literature Review\'!K5:K43,"Reading")',
     "Papers currently in progress"),
    ("% read or cited",
     '=IFERROR((COUNTIF(\'Literature Review\'!K5:K43,"Read")+COUNTIF(\'Literature Review\'!K5:K43,"Cited"))/MAX(1,COUNTA(\'Literature Review\'!B5:B43)),0)',
     "Papers read or cited ÷ total logged — use to demonstrate progress to supervisors"),
]
for i, (label, formula, note) in enumerate(lit_progress):
    ar += 1
    bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    ws_db.row_dimensions[ar].height = 16
    for ci in range(1, 9):
        c = ws_db.cell(row=ar, column=ci)
        c.fill = fill(bg); c.border = thin_border()
        c.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")
        c.font = Font(name="Calibri", size=10)
    ws_db.cell(row=ar, column=1).value = label
    ws_db.cell(row=ar, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c2 = ws_db.cell(row=ar, column=2, value=formula)
    c2.font = Font(name="Calibri", bold=True, size=10)
    c2.number_format = "0%" if label == "% read" else "0"
    c8 = ws_db.cell(row=ar, column=8, value=note)
    c8.font = Font(name="Calibri", size=9, color="64748B", italic=True)

print("✓ Dashboard built")

# ═══════════════════════════════════════════════════════════════════════════════
# GATES
# ═══════════════════════════════════════════════════════════════════════════════
ws_gt = wb.create_sheet("Gates")
ws_gt.sheet_view.showGridLines = False
ws_gt.column_dimensions["A"].width = 14
ws_gt.column_dimensions["B"].width = 16
ws_gt.column_dimensions["C"].width = 16
ws_gt.column_dimensions["D"].width = 24
ws_gt.column_dimensions["E"].width = 60
ws_gt.column_dimensions["F"].width = 16

ws_gt.row_dimensions[1].height = 26
ws_gt.merge_cells("A1:F1")
c = ws_gt["A1"]; c.value = "The four programme gates"
c.fill = fill(TITLE_BG); c.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_gt.row_dimensions[2].height = 16
ws_gt.merge_cells("A2:F2")
c = ws_gt["A2"]
c.value = "Nothing downstream of a gate formally begins until the gate condition is met."
c.fill = fill(TITLE_BG); c.font = Font(name="Calibri", size=10, color="94A3B8")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

ws_gt.row_dimensions[3].height = 18
for ci, hdr in enumerate(["Gate","Planned","Actual","Met by / condition","What it releases","Status"], 1):
    c = ws_gt.cell(row=3, column=ci, value=hdr)
    c.fill = fill(COL_BG); c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin_border()

gates_data = [
    ("GATE 1", date(2027,1,18), None,
     "Official registration confirmed – University of Essex",
     "Registration is active and student status granted. Research activity begins – including ethics applications, literature work, Qualtrics builds, and all study design. Nothing begins before this date.",
     "Not met"),
    ("GATE 2", date(2027,3,15), None,
     "UESS ethics approval received",
     "Full ethics clearance from the University of Essex Ethics Review Process. Pilot study data collection may begin. If the application is delayed this date moves and compresses everything downstream.",
     "Not met"),
    ("GATE 3", date(2027,9,30), None,
     "Annual Progress Review passed – Graduate Director and panel",
     "APR panel confirms satisfactory progress and continuation to the examination stage. All research activities already under way continue; formal submission preparation confirmed.",
     "Not met"),
    ("GATE 4", date(2028,1,18), None,
     "Marie Juanchich and Miro Sirota pre-submission sign-off",
     "Thesis formatted and submitted to the University of Essex for examination. Registration period ends. This date cannot be extended by agreement alone. Viva preparation begins.",
     "Not met"),
    ("FINAL",  date(2028,6,30), None,
     "Corrections accepted – University of Essex formal award",
     "Thesis corrected and accepted following viva voce examination. PhD (thesis incorporating publications) formally awarded. Thesis deposited in Essex Research Online and British Library EThOS.",
     "Not met"),
]

for i, (gate, planned, actual, met_by, releases, status) in enumerate(gates_data):
    r = 4 + i
    ws_gt.row_dimensions[r].height = 55
    is_final = (gate == "FINAL")
    gbg = FINL_BG if is_final else GATE_BG; gfg = FINL_FG if is_final else GATE_FG
    row_bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    vals = [gate, planned, actual if actual else None, met_by, releases, status]
    for ci, val in enumerate(vals, 1):
        c = ws_gt.cell(row=r, column=ci, value=val)
        if ci == 1:
            c.fill = fill(gbg); c.font = Font(name="Calibri", bold=True, size=10, color=gfg)
        else:
            c.fill = fill(EDIT_BG if ci in (3, 6) else row_bg)
            c.font = Font(name="Calibri", size=10, color="1E293B")
        if ci == 2 and isinstance(val, date): c.number_format = DATE_FMT
        if ci == 5:
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        else:
            c.alignment = Alignment(horizontal="center" if ci not in (1,4) else "left",
                                    vertical="center", indent=1 if ci in (1,4) else 0)
        c.border = thin_border()

print("✓ Gates built")

# ═══════════════════════════════════════════════════════════════════════════════
# DELIVERABLES
# ═══════════════════════════════════════════════════════════════════════════════
ws_dl = wb.create_sheet("Deliverables")
ws_dl.sheet_view.showGridLines = False
for ci, w in enumerate([18, 46, 20, 30, 16, 30], 1):
    ws_dl.column_dimensions[get_column_letter(ci)].width = w

ws_dl.row_dimensions[1].height = 26
ws_dl.merge_cells("A1:F1")
c = ws_dl["A1"]; c.value = "Thesis deliverable set"
c.fill = fill(TITLE_BG); c.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_dl.row_dimensions[2].height = 16
ws_dl.merge_cells("A2:F2")
c = ws_dl["A2"]
c.value = "The outputs the programme produces, and where each lands in the workplan."
c.fill = fill(TITLE_BG); c.font = Font(name="Calibri", size=10, color="94A3B8")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

ws_dl.row_dimensions[3].height = 18
for ci, hdr in enumerate(["Ref","Deliverable","Format","When in the workplan","Status","Notes"], 1):
    c = ws_dl.cell(row=3, column=ci, value=hdr)
    c.fill = fill(COL_BG); c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin_border()

deliverables = [
    ("Regulations",  "Written guidance from Essex Doctoral School on including co-authored work in standard PhD thesis (A22)",  "Written record","A-block, by Jan 2027","Not started"),
    ("Regulations",  "Agreed authorship and contribution framework with Marie and Miro (A24)",  "Written record","A-block, by Jan 2027","Not started"),
    ("Ethics pathway","Ethics pathway decision (exempt or full review)","Decision record","A-block lead-in, Oct–Nov 2026","Not started"),
    ("Ethics app",   "Pilot study ethics application (UESS)","Application","Submitted Jan 2027; Gate 2 Mar 2027","Not started"),
    ("Ethics app",   "Study 1 ethics application or amendment (UESS)","Application","Submitted Apr 2027","Not started"),
    ("Ethics app",   "Study 2 ethics application (UESS)","Application","Submitted Jul 2027","Not started"),
    ("Budget",       "Participant recruitment budget – confirmed funding source for ~1,950 Prolific participants","Decision record","B15, Feb 2027","Not started"),
    ("Pilot design", "Pilot study design document, measures, and questionnaire","Design docs","Complete by registration (Jan 2027)","Not started"),
    ("Pilot design", "Pilot study PIS and consent form","Ethics docs","With UESS application Jan 2027","Not started"),
    ("Pilot design", "Pilot study OSF pre-registration (B10 draft; C02 submitted)","Pre-reg","Drafted Feb 2027; registered Mar 2027","Not started"),
    ("Pilot design", "Pilot study Qualtrics survey (live and tested, including C13 test run)","Survey","Ready to launch Mar 2027","Not started"),
    ("Pilot write-up","Pilot study analysis, findings summary, and decision memo","Report","Apr 2027","Not started"),
    ("Study 1",      "Study 1 Statistical Analysis Plan (D09)","SAP","Apr 2027","Not started"),
    ("Study 1",      "Study 1 design document, measures, power analysis","Design docs","Finalised Apr 2027","Not started"),
    ("Study 1",      "Study 1 PIS and consent form","Ethics docs","With ethics amendment Apr 2027","Not started"),
    ("Study 1",      "Study 1 OSF pre-registration","Pre-reg","Registered before data collection May 2027","Not started"),
    ("Study 1",      "Study 1 dataset – cleaned, coded, and open data package on OSF (F19)","Dataset + package","Aug 2027","Not started"),
    ("Study 2",      "Study 2 Statistical Analysis Plan (F09)","SAP","Jul 2027","Not started"),
    ("Study 2",      "Study 2 design document and measures","Design docs","Finalised Jul 2027","Not started"),
    ("Study 2",      "Study 2 PIS and consent form","Ethics docs","With UESS application Jul 2027","Not started"),
    ("Study 2",      "Study 2 OSF pre-registration","Pre-reg","Registered before data collection Aug 2027","Not started"),
    ("Study 2",      "Study 2 dataset – cleaned, coded, and open data package on OSF (G13)","Dataset + package","Oct 2027","Not started"),
    ("Lit review",   "Systematic literature review notes, all streams","Synthesis notes","Complete by registration Jan 2027","Not started"),
    ("CTM paper",    "CTM position paper (Steps 2–5 revision) – submitted","Manuscript","Submitted Jan 2027","Not started"),
    ("Chapter 1",    "Chapter 1 – theoretical framework (final draft)","Thesis chapter","Final draft May 2027","Not started"),
    ("Chapter 2",    "Chapter 2 – empirical overview and methods","Thesis chapter","Draft complete Aug 2027","Not started"),
    ("Chapter 3",    "Chapter 3 – discussion and theoretical integration","Thesis chapter","First draft Oct 2027; revised Nov 2027","Not started"),
    ("Chapter 4",    "Chapter 4 – conclusions and contribution to knowledge","Thesis chapter","First draft Nov 2027","Not started"),
    ("Thesis",       "Thesis abstract (800–1,000 words) – H07","Thesis component","Dec 2027","Not started"),
    ("Regulations",  "Candidate's contribution statement (signed by co-authors) for each co-authored paper – H08","Declaration","Dec 2027","Not started"),
    ("Paper 1",      "Paper 1 – Study 1 findings – cover letter (G05) + submission (G06)","Manuscript","Draft Aug 2027; submitted Oct 2027","Not started"),
    ("Paper 2",      "Paper 2 – Study 2 findings – cover letter (H02) + submission (H03)","Manuscript","Draft Oct 2027; submitted Nov 2027","Not started"),
    ("Conference",   "Conference abstract (APA, BPS, SPSP, or equivalent) and presentation","Abstract + talk","Abstract May 2027; presentation varies","Not started"),
    ("APR",          "Annual Progress Review written report","Report","Submitted Sep 2027; Gate 3","Not started"),
    ("Thesis",       "Full thesis for examination (all chapters integrated)","Thesis","Full draft Dec 2027; submitted Jan 2028","Not started"),
    ("Thesis",       "Thesis deposited in Essex Research Online (ERO) and British Library EThOS – J10","Repository deposit","Jun 2028 (post-award)","Not started"),
]

dv_dl = DataValidation(type="list",
    formula1='"Not started,In progress,Complete,Submitted,Accepted,Rejected – revising"',
    allow_blank=True, showDropDown=False)
dv_dl.sqref = f"E4:E{3+len(deliverables)}"
ws_dl.add_data_validation(dv_dl)

for i, row_data in enumerate(deliverables):
    r = 4 + i
    ws_dl.row_dimensions[r].height = 22
    bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
    ref, deliv, fmt, when, status = row_data
    for ci, val in enumerate([ref, deliv, fmt, when, status, ""], 1):
        c = ws_dl.cell(row=r, column=ci, value=val)
        c.fill = fill(EDIT_BG if ci == 5 else bg)
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="left" if ci in (2,4,6) else "center",
                                vertical="center", wrap_text=True, indent=1)
        c.border = thin_border()

r = 4 + len(deliverables)
ws_dl.row_dimensions[r].height = 17
for ci in range(1, 7):
    c = ws_dl.cell(row=r, column=ci)
    c.fill = fill(TOTL_BG); c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    c.border = thin_border()
    c.alignment = Alignment(horizontal="center" if ci>1 else "left", vertical="center")
ws_dl.cell(row=r, column=1).value = "Delivered"
ws_dl.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_dl.cell(row=r, column=5).value = (f'=COUNTIF(E4:E{r-1},"Complete")'
                                      f'&" of {len(deliverables)}"')

print("✓ Deliverables built")

# ═══════════════════════════════════════════════════════════════════════════════
# LITERATURE REVIEW
# ═══════════════════════════════════════════════════════════════════════════════
ws_lr = wb.create_sheet("Literature Review")
ws_lr.sheet_view.showGridLines = False

LR_HDRS = ["Ref #", "Author(s)", "Year", "Title", "Journal / Source",
           "DOI / URL", "Keywords used to find", "Stream",
           "Key argument / finding", "Relevance to thesis", "Status"]
LR_COL_WIDTHS = [9, 28, 7, 40, 26, 32, 30, 22, 42, 38, 14]
NC_LR = len(LR_HDRS)  # 11
last_lr = get_column_letter(NC_LR)  # K

for i, w in enumerate(LR_COL_WIDTHS, 1):
    ws_lr.column_dimensions[get_column_letter(i)].width = w

# ── Row 1: Title banner ───────────────────────────────────────────────────────
ws_lr.row_dimensions[1].height = 26
ws_lr.merge_cells(f"A1:{last_lr}1")
c = ws_lr["A1"]
c.value = "Literature Review Log  ·  PhD by Publication"
c.fill = fill(LR_BG)
c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── Row 2: Subtitle / instructions ───────────────────────────────────────────
ws_lr.row_dimensions[2].height = 18
ws_lr.merge_cells(f"A2:{last_lr}2")
c = ws_lr["A2"]
c.value = ("Geoff Whittle-Walls  ·  Track every paper, its stream, and where it will be cited  ·  "
           "Status column = To read / Reading / Read / Cited  ·  "
           "Shaded cells are editable  ·  Add keyword searches in the lower table")
c.fill = fill(LR_BG)
c.font = Font(name="Calibri", size=9, color="A7F3D0")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── Row 3: Section 1 header ───────────────────────────────────────────────────
ws_lr.row_dimensions[3].height = 18
ws_lr.merge_cells(f"A3:{last_lr}3")
c = ws_lr["A3"]
c.value = "SECTION 1  ·  CITATION LOG  ·  Add one row per paper; sort or filter by Stream or Status"
c.fill = fill(LR_SECT)
c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── Row 4: Column headers ─────────────────────────────────────────────────────
ws_lr.row_dimensions[4].height = 30
hdr_fill = fill("1E3A5F")
for col_idx, h in enumerate(LR_HDRS, 1):
    c = ws_lr.cell(row=4, column=col_idx, value=h)
    c.fill = hdr_fill
    c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    c.border = thin_border()

ws_lr.freeze_panes = "A5"

# ── Paper data (LR01–LR23) ────────────────────────────────────────────────────
LR_PAPERS = [
    ("LR01", "Windschitl, P. D. & Wells, G. L.", 1996,
     "Measuring psychological uncertainty: Verbal versus numeric methods",
     "Journal of Experimental Psychology: Applied",
     "https://doi.org/10.1037/1076-898X.2.4.343",
     "verbal probability; numeric probability; uncertainty; measurement; format",
     "Uncertainty communication",
     "Verbal and numeric probability expressions activate different cognitive processes and elicit different "
     "responses. Verbal forms are more ambiguous but carry evaluative and emotional weight; numeric formats "
     "convey precision but lack affective nuance. Sender–receiver mapping is systematically inconsistent "
     "across formats.",
     "Foundational for Study 1 and Study 2 stimulus design: establishes that format (verbal vs numeric) is not "
     "a neutral carrier of information. Grounds the thesis claim that how uncertainty is expressed shapes "
     "perception independently of the underlying probability value.",
     "To read"),
    ("LR02", "Teigen, K. H.", 2023,
     "Dimensions of uncertainty communication: What is conveyed by verbal terms and numeric ranges",
     "Current Psychology, 42, 29122–29137",
     "https://doi.org/10.1007/s12144-022-03985-0",
     "uncertainty; verbal probability; numeric ranges; communication; interpretation",
     "Uncertainty communication",
     "Compares verbal uncertainty expressions and numeric ranges across multiple dimensions: directionality, "
     "precision, and evaluative tone. Shows that verbal terms carry implicit directional and evaluative meaning "
     "beyond mere probability, while ranges signal precision. Receivers systematically interpret both forms "
     "differently from senders’ intent. Recommends matching format to communicative goal.",
     "Directly informs Study 1 and Study 2 design: the most recent treatment of verbal vs numeric uncertainty "
     "and its interpretation gaps. Chapter 1 theoretical framework on uncertainty communication rests "
     "substantially on this paper.",
     "To read"),
    ("LR03", "Fischhoff, B.", 1995,
     "Risk perception and communication unplugged: Twenty years of process",
     "Risk Analysis, 15(2), 137–145",
     "https://doi.org/10.1111/j.1539-6924.1995.tb00308.x",
     "risk communication; mental models; audience; process; expert-public gap",
     "Uncertainty communication",
     "Reviews 20 years of risk communication research and practice. Argues that effective risk communication "
     "requires understanding the audience’s existing mental models and systematically filling knowledge "
     "gaps. Identifies structural reasons why one-way information provision fails; advocates for iterative, "
     "audience-centred communication processes.",
     "Frames the core problem the PhD addresses: why accurate risk/uncertainty communication fails to change "
     "beliefs or behaviour. Mental models approach directly informs Study 1 design and the theoretical "
     "discussion in Chapter 1.",
     "To read"),
    ("LR04", "Slovic, P.", 1987,
     "Perception of risk",
     "Science, 236(4799), 280–285",
     "https://doi.org/10.1126/science.3563507",
     "risk perception; dread risk; psychometric paradigm; laypeople; experts",
     "Risk perception",
     "Established the psychometric paradigm for risk perception: identified two dominant factors (dread risk "
     "and unknown risk) that explain variance in perceived risk far beyond objective probability estimates. "
     "Laypeople and experts rank risks systematically differently. Risk is multidimensional; subjective "
     "perception cannot be reduced to expected-value calculations.",
     "One of the most-cited papers in risk psychology. Foundational for Chapter 1 theoretical framework: "
     "establishes that perceived risk is a constructed, multidimensional judgment – not a simple reading "
     "of probability. Grounds Study 1 and Study 2 design by motivating why format effects on risk perception "
     "are psychologically meaningful.",
     "To read"),
    ("LR05", "Slovic, P.", 1993,
     "Perceived Risk, Trust, and Democracy",
     "Risk Analysis, 13(6), 675–682",
     "https://doi.org/10.1111/j.1539-6924.1993.tb01329.x",
     "risk perception; trust; democracy; public; asymmetry",
     "Risk perception",
     "Trust in risk managers is constructed through experience and is asymmetric: easy to destroy, hard to "
     "build. Democratic institutions must earn public trust through transparency and responsiveness. Risk "
     "perception is heavily influenced by perceived fairness and trust in the information source; this trust "
     "asymmetry has major implications for public communication strategy.",
     "Directly relevant to the thesis argument that responses to uncertainty communications depend on source "
     "trust, not content alone. Recommended by supervisor Marie. Relevant to Chapter 1 and Study 1 design.",
     "To read"),
    ("LR06", "Loewenstein, G. F., Weber, E. U., Hsee, C. K. & Welch, N.", 2001,
     "Risk as feelings",
     "Psychological Bulletin, 127(2), 267–286",
     "https://doi.org/10.1037/0033-2909.127.2.267",
     "risk; affect; feelings; emotion; dual process; decision making",
     "Risk perception",
     "Proposes the risk-as-feelings hypothesis: emotional responses to risky situations drive behaviour "
     "independently of – and sometimes contrary to – cognitive risk appraisals. Feelings anticipate "
     "consequences and react to vividness, probability neglect, and regret. Identifies systematic divergence "
     "between ‘what we think’ and ‘what we feel’ about risk.",
     "Explains why accurate uncertainty communication may fail to change behaviour when emotional responses "
     "are not addressed. Directly relevant to Study 1 and Study 2 intervention design; grounds the affect "
     "heuristic discussion in Chapter 1 theoretical framework.",
     "To read"),
    ("LR07", "Levine, T. R.", 2014,
     "Truth-Default Theory (TDT): A theory of human deception and deception detection",
     "Journal of Language and Social Psychology, 33(4), 378–392",
     "https://doi.org/10.1177/0261927X14535916",
     "truth default theory; TDT; deception; lie detection; veracity effect; trigger model",
     "Truth default theory",
     "Formal presentation of Truth-Default Theory (TDT). Core claims: humans default to assuming honesty; "
     "active lie detection is triggered only by suspicion cues or motivated circumstances; the veracity effect "
     "(truth detected better than lies on average) arises from the combination of truth default and base-rate "
     "honesty. Provides the most comprehensive theoretical statement of TDT.",
     "Primary theoretical anchor for the TDT stream. Introduces core concepts (truth default, trigger model, "
     "veracity effect) used in Chapter 1. Must be read before LR08 (Levine 2022) and LR09 (Moore & Hancock "
     "2026). Central citation for Papers 1 and 2.",
     "To read"),
    ("LR08", "Levine, T. R.", 2022,
     "Truth-default theory and the psychology of lying and deception detection",
     "Current Opinion in Psychology",
     "https://doi.org/10.1016/j.copsyc.2022.101380",
     "truth default theory; deception; lie detection; TDT; communication",
     "Truth default theory",
     "Concise state-of-science overview of TDT: the default assumption of honesty is the evolved "
     "communicative norm; lie detection is poor when passive, triggered only by suspicion cues. TDT explains "
     "why misinformation spreads. Reviews empirical support across contexts and contrasts TDT with the "
     "adaptive lie detector account.",
     "Grounds the thesis in TDT as a communication theory. Explains why audiences fail to detect misleading "
     "uncertainty framing by default. Informs the theoretical model linking uncertainty framing to perceived "
     "deception. Recommended by supervisor Marie.",
     "To read"),
    ("LR09", "Moore, R. C. & Hancock, J. T.", 2026,
     "Extending Truth-Default Theory to Misinformation: Connecting Misinformation Scholarship and Deception Research",
     "Perspectives on Psychological Science",
     "https://doi.org/10.1177/17456916261475589",
     "truth default theory; misinformation; deception; TDT; persuasion",
     "Truth default theory",
     "Bridges TDT and misinformation research: argues that the mechanisms governing interpersonal deception "
     "(truth default, trigger model, veracity effect) operate in how people evaluate mediated misinformation. "
     "Connects sender-level deception literature to audience-level information processing. Most recent "
     "synthesis at programme start (published October 2026).",
     "Critical bridge paper: links TDT (LR07, LR08) to the misinformation context. Justifies applying a "
     "deception/trust framework to uncertainty communication in health and risk contexts. Central citation "
     "for Chapter 1 and Paper 1 introduction. Recommended by supervisor Marie.",
     "To read"),
    ("LR10", "Juanchich, M., Sirota, M. & Holford, D. L.", 2023,
     "How Should Doctors Frame the Risk of a Vaccine’s Adverse Side Effects? It Depends on How Trustworthy They Are",
     "Medical Decision Making, 43(7–8), 835–849",
     "https://doi.org/10.1177/0272989X231197646",
     "vaccine; side effects; framing; source trust; vaccination intention; risk communication; utility",
     "Vaccine communication",
     "When a distrusted doctor downplays vaccine side-effect risks, it triggers reactance and lowers "
     "vaccination intention – the opposite of the intended effect. Trusted sources can safely frame risks "
     "away; distrusted sources must disclose risks transparently to preserve credibility. Experimentally "
     "demonstrates the source trust × framing direction interaction on intention.",
     "Primary supervisor paper: the applied anchor for the whole PhD. Directly cited in the proposal. "
     "Establishes the empirical phenomenon that drives the thesis – communicator format × source "
     "trust interaction on intention. Introduces the SDT utility framework: communicator criterion reflects "
     "downplaying vs over-stating. Read first in this stream before LR11 and LR12.",
     "To read"),
    ("LR11", "Juanchich, M., Oakley, C. M., Sayer, H., Holford, D. L., Bruine de Bruin, W., Booker, C., "
             "Chadborn, T., Vallee-Tourangeau, G., Wood, R. & Sirota, M.", 2024,
     "Vaccination invitations sent by warm and competent medical professionals disclosing risks and benefits "
     "increase trust and booking intention and reduce inequalities between ethnic groups",
     "Health Psychology, 43(10), 718–729",
     "https://doi.org/10.1037/hea0001385",
     "vaccination; trust; invitation; risk disclosure; benefit framing; ethnic minority; source credibility",
     "Vaccine communication",
     "Invitation letters from warm and competent GPs that explicitly disclose vaccine risks alongside benefits "
     "increase both trust and booking intention, with the largest gains in ethnic minority groups who typically "
     "distrust the source. Transparent risk-benefit disclosure reduces trust concerns rather than amplifying "
     "them. Effect of message content outweighs source warmth when both are varied.",
     "Primary supervisor paper: directly cited in proposal. Shows that describing risks relative to benefits "
     "(not in isolation) improves outcomes, especially for those who distrust the source – the key "
     "moderator in Studies 2 and 3. Reinforces the thesis argument that utility framing × baseline trust "
     "interactively shape vaccination intention and perceived communicator intent.",
     "To read"),
    ("LR12", "Thorpe, A., Fagerlin, A., Drews, F. A., Butler, J., Stevens, V., Riddoch, M. S. & Scherer, L. D.", 2022,
     "Communications to Promote Interest and Confidence in COVID-19 Vaccines",
     "American Journal of Health Promotion, 36(6), 976–986",
     "https://doi.org/10.1177/08901171221082904",
     "COVID-19; vaccine; side effects; framing; vaccination interest; confidence; contextualisation",
     "Vaccine communication",
     "In a within-subjects online survey of 1,043 US adults, side-effect information presented in isolation "
     "produced the lowest vaccine interest. Adding benefit information, reframing side-effect likelihood, or "
     "contextualising post-vaccine symptoms as immune activation signals each significantly increased vaccine "
     "interest above the isolated side-effects baseline.",
     "Directly cited in the proposal. Establishes that side-effect risks in isolation reduce vaccination "
     "intention – the control condition for Studies 2 and 3. Grounds the prediction that contextualised "
     "(relative) risk presentation outperforms isolated risk presentation, motivating the framing × "
     "format manipulations.",
     "To read"),
    ("LR13", "Sandman, P. M.", 1989,
     "Hazard versus outrage in the public perception of risk",
     "In V. T. Covello, D. B. McCallum & M. T. Pavlova (Eds.), Effective risk communication (pp. 45–49). Plenum Press",
     "https://www.psandman.com/articles/risk.htm",
     "risk perception; outrage; hazard; risk dimensions; voluntariness; dread; trust; familiarity",
     "Risk perception",
     "Proposes the Hazard + Outrage model: public risk perception is driven not only by objective hazard "
     "(probability × consequences) but by psychosocial ‘outrage’ factors including "
     "voluntariness, familiarity, trust in the source, catastrophic potential, controllability, and natural "
     "vs industrial origin. Outrage factors often dominate objective hazard in predicting public response.",
     "Directly cited in the pilot study design. Sandman’s dimensions – voluntary/non-voluntary, "
     "natural/industrial, familiar/unfamiliar, chronic/catastrophic, controllable/uncontrollable, "
     "trusted/untrusted source – are the criterion for selecting non-nuclear risk stimuli that differ "
     "systematically from nuclear risk. Grounds stimulus selection rationale in the pilot study.",
     "To read"),
    ("LR14", "Friestad, M. & Wright, P.", 1994,
     "The Persuasion Knowledge Model: How People Cope with Persuasion Attempts",
     "Journal of Consumer Research, 21(1), 1–31",
     "https://doi.org/10.1086/209380",
     "persuasion knowledge; intent inference; communicator motive; coping; scepticism",
     "Persuasion knowledge",
     "People develop lay theories about persuasion tactics and deploy this knowledge to infer communicator "
     "intent and adjust their responses. Persuasion knowledge is activated when a source is seen as trying "
     "to change beliefs; it drives scepticism, counter-arguing, and avoidance. Perceived motive attribution "
     "is shaped more by format and tactic recognition than by message content alone.",
     "Foundational for the thesis: the Persuasion Knowledge Model explains how audiences infer whether a "
     "communicator is trying to downplay or overstate risk based on the format they choose. Central to the "
     "intention inference framework in Chapter 1 and the theoretical rationale for RQ1 (What drives "
     "intention perception?) and RQ2 (How do perceived utilities shape it?).",
     "To read"),
    ("LR15", "Evans, J. St. B. T.", 2008,
     "Dual-processing accounts of reasoning, judgment, and social cognition",
     "Annual Review of Psychology, 59, 255–278",
     "https://doi.org/10.1146/annurev.psych.59.103006.093629",
     "dual process theory; System 1; System 2; intuition; analytic reasoning; judgment",
     "Dual process theory",
     "Comprehensive review of System 1 (fast, automatic, intuitive) vs System 2 (slow, deliberate, analytic) "
     "processing across reasoning, judgment, and social cognition. Synthesises evidence supporting the "
     "two-system framework and clarifies theoretical disputes (parallel vs sequential processing, "
     "default-interventionist vs parallel-competitive models).",
     "Grounds the CTM position paper and uncertainty communication framework: explains when intuitive vs "
     "deliberate processing drives responses to risk and uncertainty communications. Chapter 1 theoretical "
     "synthesis; relevant to individual differences moderating Study 1 and Study 2 effects.",
     "To read"),
    ("LR16", "Pennycook, G. & Rand, D. G.", 2019,
     "Lazy, not biased: Susceptibility to partisan fake news is better explained by lack of reasoning than "
     "by motivated reasoning",
     "Cognition, 188, 39–50",
     "https://doi.org/10.1016/j.cognition.2018.06.011",
     "analytical thinking; fake news; misinformation; motivated reasoning; numeracy",
     "Dual process theory",
     "Individual differences in analytical thinking (CRT) predict fake news discernment across partisan "
     "lines: susceptibility is driven by lazy thinking rather than political motivated reasoning. High-CRT "
     "participants reject partisan misinformation more than low-CRT participants regardless of political "
     "alignment. Implications for interventions targeting analytic engagement.",
     "Informs how individual differences in reasoning style moderate responses to uncertainty communications "
     "in Study 1 and Study 2. Supports the thesis argument that dual-process engagement moderates whether "
     "people detect misleading uncertainty framing. Chapter 1 and Papers 1 and 2.",
     "To read"),
    ("LR17", "Peters, E., Slovic, P., Västfjäll, D. & Mertz, C. K.", 2008,
     "Intuitive numbers guide decisions",
     "Judgment and Decision Making, 3(8), 619–635",
     "https://doi.org/10.1037/h0087705",
     "numeracy; affect; decision making; framing; risk communication; individual differences",
     "Numeracy",
     "High-numeracy individuals extract more precise affective meaning from numbers and this drives better "
     "decisions under risk. Low-numeracy individuals are more susceptible to affective and framing influences "
     "on numeric information. Numeracy moderates the relationship between numerical risk information and "
     "decision quality.",
     "Directly relevant to Study 1 and Study 2 design: numeracy is a key individual difference moderating "
     "how people respond to numerical uncertainty communications. Grounds the thesis prediction that numeracy "
     "interacts with uncertainty format to affect comprehension and risk perception.",
     "To read"),
    ("LR18", "Tversky, A. & Kahneman, D.", 1981,
     "The framing of decisions and the psychology of choice",
     "Science, 211(4481), 453–458",
     "https://doi.org/10.1126/science.7455683",
     "framing; loss aversion; prospect theory; decision making; Asian disease problem",
     "Framing effects",
     "Demonstrated that logically equivalent options framed as gains vs losses produce systematically "
     "different choices, violating rational choice theory. The Asian Disease Problem became the canonical "
     "demonstration of framing effects. Risk-seeking in loss frames and risk-aversion in gain frames are "
     "robust and replicable across contexts.",
     "Core citation for the framing stream. Establishes that how uncertainty is presented (not just what is "
     "presented) systematically alters decisions. Grounds Study 1 and Study 2 design manipulations; Chapter "
     "1 theoretical framework on format effects in risk communication.",
     "To read"),
    ("LR19", "Levin, I. P., Schneider, S. L. & Gaeth, G. J.", 1998,
     "All frames are not created equal: A typology and critical analysis of framing effects",
     "Organizational Behavior and Human Decision Processes, 76(2), 149–188",
     "https://doi.org/10.1006/obhd.1998.2804",
     "framing; typology; risky choice framing; attribute framing; goal framing",
     "Framing effects",
     "Proposed a three-type taxonomy of framing effects: risky choice framing (Tversky & Kahneman), "
     "attribute framing (single attribute as positive/negative), and goal framing (emphasising gain or loss "
     "of engaging in behaviour). Each type has different mechanisms and moderators. Argues that ‘framing’ "
     "is not a single phenomenon.",
     "Provides conceptual precision for Study 1 and Study 2 design: uncertainty framing likely involves "
     "attribute framing. Helps Geoff specify which type of framing effect is predicted and measured. "
     "Chapter 1 theoretical framework and Papers 1 and 2 Methods sections.",
     "To read"),
    ("LR20", "Thagard, P.", 1989,
     "Explanatory coherence",
     "Behavioral and Brain Sciences, 12(3), 435–467",
     "https://doi.org/10.1017/S0140525X00057046",
     "explanatory coherence; ECHO; constraint satisfaction; belief revision; reasoning",
     "Coherence theory",
     "Proposes a computational theory of coherence in scientific reasoning. Beliefs are accepted or rejected "
     "through constraint satisfaction maximising overall coherence across a network of propositions. The ECHO "
     "model demonstrates that coherence principles predict theory choice better than simple induction or "
     "confirmation logic.",
     "Foundational for Geoff’s CTM position paper and the coherence-based framework for understanding "
     "how people integrate uncertain or conflicting information. Chapter 1 theoretical synthesis; defines "
     "the coherence mechanisms that the thesis proposes underpin responses to uncertainty communication.",
     "To read"),
    ("LR21", "Grice, H. P.", 1975,
     "Logic and conversation",
     "In P. Cole & J. Morgan (Eds.), Syntax and semantics, Vol. 3: Speech acts (pp. 41–58). Academic Press",
     "https://doi.org/10.1163/9789004368811_003",
     "implicature; cooperative principle; Gricean maxims; pragmatics; speaker intent; quantity; manner",
     "Communicative pragmatics",
     "Proposes the cooperative principle and four conversational maxims (quantity, quality, relation, manner). "
     "When speakers appear to violate a maxim, listeners generate conversational implicatures – additional "
     "meaning beyond the literal content – to preserve the assumption of cooperation. Speakers signal "
     "intent not only through what they say, but through format choices and what they omit.",
     "Provides the linguistic and pragmatic foundation for the thesis: format choice (verbal vs numerical, "
     "focusing vs downplaying risks) signals communicator intent through Gricean implicature. When a doctor "
     "uses a vague verbal form for a vaccine risk, the audience uses quantity and manner implicatures to infer "
     "motivation. Chapter 1 theoretical framework; grounds RQ1 and RQ3.",
     "To read"),
    ("LR22", "Kahneman, D.", 2003,
     "A perspective on judgment and choice: Mapping bounded rationality",
     "American Psychologist, 58(9), 697–720",
     "https://doi.org/10.1037/0003-066X.58.9.697",
     "bounded rationality; System 1; System 2; heuristics; biases; Nobel lecture",
     "Dual process theory",
     "Nobel Prize lecture synthesising decades of research on heuristics and biases. Articulates the "
     "two-system framework most clearly for a general scientific audience. Reviews the accessibility "
     "principle, affective heuristic, and their roles in judgment under uncertainty. Identifies the "
     "conditions under which System 2 overrides System 1 errors.",
     "Most accessible statement of the dual-process framework foundational to the thesis. Should be read "
     "after LR15 (Evans 2008). Chapter 1 theoretical framework; provides the individual differences context "
     "for Studies 1 and 2.",
     "To read"),
    ("LR23", "Gigerenzen, G. & Edwards, A.", 2003,
     "Simple tools for understanding risks: From innumeracy to insight",
     "BMJ, 327(7417), 741–744",
     "https://doi.org/10.1136/bmj.327.7417.741",
     "natural frequencies; risk communication; numeracy; health; icon arrays",
     "Numeracy",
     "Demonstrates that presenting risks as natural frequencies (e.g., ‘10 in 1,000’) rather than "
     "probabilities (e.g., ‘1%’) dramatically improves comprehension across numeracy levels. "
     "Proposes simple tools (icon arrays, natural frequencies, absolute rather than relative risk) to improve "
     "health risk communication. Highly cited across medicine and psychology.",
     "Directly informs Study 1 uncertainty stimulus design: natural frequency format vs probability format "
     "as a key experimental manipulation. Chapter 1 and Papers 1 and 2 Methods sections. Provides applied "
     "relevance for the thesis.",
     "To read"),
]

# ── Populate rows 5–27 (LR01–LR23) ───────────────────────────────────────────
for row_idx, paper in enumerate(LR_PAPERS, 5):
    ref, authors, year, title, journal, doi, keywords, stream, key_arg, relevance, status = paper
    row_data = [ref, authors, year, title, journal, doi, keywords, stream, key_arg, relevance, status]
    row_fill = fill(ROW_EVEN) if row_idx % 2 == 0 else fill(ROW_ODD)
    for col_idx, val in enumerate(row_data, 1):
        c = ws_lr.cell(row=row_idx, column=col_idx, value=val)
        c.border = thin_border()
        if col_idx == 11:  # Status – editable highlight
            c.fill = fill(EDIT_BG)
        else:
            c.fill = row_fill
        c.font = cell_font(size=9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws_lr.row_dimensions[row_idx].height = 45

# ── Blank rows 28–43 (LR24–LR39) ─────────────────────────────────────────────
for blank_num in range(24, 40):
    row_idx = blank_num - 24 + 28  # rows 28–43
    ref_label = f"LR{blank_num:02d}"
    row_fill = fill(ROW_EVEN) if row_idx % 2 == 0 else fill(ROW_ODD)
    c = ws_lr.cell(row=row_idx, column=1, value=ref_label)
    c.fill = row_fill
    c.font = Font(name="Calibri", size=9, color="94A3B8")  # dimmed grey
    c.border = thin_border()
    c.alignment = Alignment(vertical="top")
    for col_idx in range(2, 12):
        c = ws_lr.cell(row=row_idx, column=col_idx)
        c.fill = fill(EDIT_BG) if col_idx == 11 else row_fill
        c.border = thin_border()
    ws_lr.row_dimensions[row_idx].height = 20

# ── Row 44: Summary row ───────────────────────────────────────────────────────
lr_data_end = 43
ws_lr.row_dimensions[44].height = 18
for col_idx in range(1, 12):
    c = ws_lr.cell(row=44, column=col_idx)
    c.fill = fill(LR_SECT)
    c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border()
ws_lr["A44"].value = "Papers logged"
ws_lr["A44"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_lr["F44"].value = f"=COUNTA(B5:B{lr_data_end})"
ws_lr["G44"].value = f'=COUNTIF(K5:K{lr_data_end},"To read")'
ws_lr["H44"].value = f'=COUNTIF(K5:K{lr_data_end},"Reading")'
ws_lr["I44"].value = f'=COUNTIF(K5:K{lr_data_end},"Read")'
ws_lr["J44"].value = f'=COUNTIF(K5:K{lr_data_end},"Cited")'

# ── Row 46: Section 2 header ──────────────────────────────────────────────────
ws_lr.row_dimensions[46].height = 18
ws_lr.merge_cells(f"A46:{last_lr}46")
c = ws_lr["A46"]
c.value = ("SECTION 2  ·  KEYWORD SEARCH LOG  ·  "
           "Document every search string and database – this becomes part of your methodology")
c.fill = fill(KW_BG)
c.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

# ── Row 47: Instructions ──────────────────────────────────────────────────────
ws_lr.row_dimensions[47].height = 50
ws_lr.merge_cells(f"A47:{last_lr}47")
c = ws_lr["A47"]
c.value = (
    "Why this matters: Marie’s advice is to keep track of every keyword combination you use and what "
    "it returns. This log IS your literature search methodology – it demonstrates rigour and "
    "reproducibility. Include it in your thesis as an appendix or Chapter 1 methods section. Also try "
    "forward citation search: find who cites the seed papers (LR01–LR23 above) to capture the most "
    "current work on the same topic. Two forward citation seed searches are already pre-populated below."
)
c.fill = fill("EFF6FF")
c.font = cell_font(size=9)
c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)

# ── Row 48: Search log headers ────────────────────────────────────────────────
KW_HDRS = ["Search date", "Database / Source", "Search string / keywords", "Stream",
           "# results", "# screened titles", "# read abstracts", "# retained",
           "Forward citation of", "Notes / next steps"]
ws_lr.row_dimensions[48].height = 30
for col_idx, h in enumerate(KW_HDRS, 1):
    c = ws_lr.cell(row=48, column=col_idx, value=h)
    c.fill = fill(KW_BG)
    c.font = Font(name="Calibri", bold=True, size=9, color="FFFFFF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    c.border = thin_border()

# ── Rows 49–68: 20 keyword search entries ─────────────────────────────────────
KW_SEARCHES = [
    ("Oct 2026", "PsycINFO",
     '("uncertainty communication" OR "verbal probability" OR "numeric probability") AND (interpretation OR comprehension OR perception)',
     "Uncertainty communication", None, None, None, None, None,
     "Seed search for Stream 1. Filter to empirical studies post-2000. Adjust OR terms if results thin."),
    ("Oct 2026", "Web of Science",
     'TS=("risk communication" OR "uncertainty expression") AND TS=("format" OR "framing" OR "verbal" OR "numeric") AND TS=(decision* OR judgment*)',
     "Uncertainty communication", None, None, None, None, None,
     "Broader Web of Science search for uncertainty format effects literature."),
    ("Oct 2026", "PsycINFO",
     '("risk perception" OR "perceived risk") AND (trust OR communication OR "mental model") AND NOT "financial risk"',
     "Risk perception", None, None, None, None, "LR04 Slovic 1987 / LR05 Slovic 1993",
     "Forward citation search from Slovic 1987 and 1993 to capture post-2010 updates."),
    ("Nov 2026", "PubMed",
     '("risk perception" OR "risk communication") AND (health OR medical) AND (numeracy OR comprehension OR "lay public")',
     "Risk perception / numeracy", None, None, None, None, None,
     "Target health-domain risk communication for Study 1 applied context."),
    ("Oct 2026", "Communication Abstracts / PsycINFO",
     '"truth default theory" OR "TDT" AND (Levine OR deception OR "lie detection" OR misinformation)',
     "Truth default theory", None, None, None, None, "LR07 Levine 2014 / LR08 Levine 2022",
     'TDT literature is concentrated around Levine; also search "truth-default" hyphenated. Check cites of Levine 2014.'),
    ("Oct 2026", "Google Scholar",
     '"truth default theory" OR "truth default" AND misinformation',
     "Truth default theory", None, None, None, None, "LR09 Moore & Hancock 2026",
     "Forward citation search from Moore & Hancock 2026 – most recent TDT+misinformation bridge."),
    ("Nov 2026", "PsycINFO",
     '("vaccination intention" OR "vaccine hesitancy" OR "vaccine acceptance") AND ("side effects" OR "adverse effects") AND (framing OR format OR "risk communication")',
     "Vaccine communication", None, None, None, None,
     "LR10 Juanchich 2023 / LR11 Juanchich 2024 / LR12 Thorpe 2022",
     "Forward search from supervisor papers (Juanchich 2023, 2024) and Thorpe 2022. Filter to experimental "
     "studies post-2015. Adjust to include ‘benefit framing’ OR ‘contextualised risk’."),
    ("Nov 2026", "PubMed",
     '("risk framing" OR "attribute framing" OR "comparative risk") AND (vaccine OR vaccination) AND (trust OR intention OR confidence OR "source credibility")',
     "Vaccine communication", None, None, None, None, "LR10 Juanchich 2023",
     "Target vaccine health communication: comparative vs isolated risk presentations. Also try "
     "‘relative risk’ AND ‘vaccination’ AND psychology."),
    ("Nov 2026", "PsycINFO",
     '("persuasion knowledge" OR "intent inference" OR "communicator intent" OR "motive attribution") AND (risk OR health OR communication OR framing)',
     "Persuasion knowledge", None, None, None, None, "LR14 Friestad & Wright 1994",
     "Forward cite from Friestad & Wright 1994 (>2000 citations). Filter to risk/health contexts. "
     "Also search ‘perceived intent’ AND ‘health communication’."),
    ("Nov 2026", "Web of Science",
     'TS=("persuasion knowledge" OR "coping with persuasion") AND TS=("source credibility" OR trust OR scepticism OR "motivated scepticism")',
     "Persuasion knowledge", None, None, None, None, "LR14 Friestad & Wright 1994",
     "Identify moderators of persuasion knowledge activation relevant to RQ4 (baseline trust shapes "
     "intention perception)."),
    ("Nov 2026", "PsycINFO",
     '("dual process" OR "System 1" OR "System 2" OR "Type 1" OR "Type 2") AND (reasoning OR judgment OR "decision making")',
     "Dual process theory", None, None, None, None, "LR15 Evans 2008 / LR22 Kahneman 2003",
     "Foundational stream – prioritise review articles. Also search ‘dual-process’ hyphenated."),
    ("Nov 2026", "PsycINFO",
     '"cognitive reflection" OR "CRT" AND (misinformation OR "fake news" OR analytical thinking)',
     "Dual process theory", None, None, None, None, "LR16 Pennycook & Rand 2019",
     "Individual differences in analytical thinking as moderator. Forward cite Pennycook & Rand 2019."),
    ("Nov 2026", "PubMed",
     '(numeracy OR "risk literacy" OR "statistical literacy") AND ("risk communication" OR health OR comprehension OR decision)',
     "Numeracy", None, None, None, None, "LR17 Peters et al 2008 / LR23 Gigerenzen & Edwards 2003",
     "Health numeracy literature. Also try ‘innumeracy’ and ‘natural frequencies’ as specific search terms."),
    ("Nov 2026", "Web of Science",
     'TS=("natural frequencies" OR "icon array" OR "frequency format") AND TS=("risk communication" OR comprehension OR decision)',
     "Numeracy", None, None, None, None, "LR23 Gigerenzen & Edwards 2003",
     "Presentation format literature; identifies papers on natural frequency vs probability format."),
    ("Dec 2026", "PsycINFO",
     '("framing effect" OR "gain frame" OR "loss frame") AND (risk OR health OR "decision making")',
     "Framing effects", None, None, None, None, "LR18 Tversky & Kahneman 1981 / LR19 Levin et al 1998",
     "Classic stream – scope to post-2000 empirical studies. Distinguish risky choice vs attribute "
     "framing (Levin 1998 taxonomy)."),
    ("Dec 2026", "PsycINFO / Google Scholar",
     '("explanatory coherence" OR "coherence theory" OR "constraint satisfaction") AND (reasoning OR belief OR judgment OR cognition)',
     "Coherence theory", None, None, None, None, "LR20 Thagard 1989",
     "Thagard’s ECHO model and applications. Niche stream – search also ‘CTM misinformation’ "
     "and ‘coherence belief revision’."),
    ("Dec 2026", "PsycINFO / Linguistics databases",
     '("conversational implicature" OR "Gricean maxim" OR "pragmatic inference" OR "cooperative principle") AND ("risk communication" OR "health communication" OR intent OR motive)',
     "Communicative pragmatics", None, None, None, None, "LR21 Grice 1975",
     "Grice’s implicature applied to health/risk communication. Also search ‘verbal format’ AND "
     "‘speaker intent’ AND risk. Consider linguistics journals (Pragmatics, Journal of Pragmatics)."),
    ("Jan 2027", "Google Scholar",
     "Cites: Juanchich, Sirota & Holford (2023) Medical Decision Making 43(7-8) 835-849",
     "Vaccine communication", None, None, None, None, "LR10 Juanchich et al 2023",
     "Forward citation pass – capture post-2023 replications and extensions of the framing × "
     "trust interaction central to the thesis."),
    ("Jan 2027", "Google Scholar",
     "Cites: Friestad & Wright (1994) Journal of Consumer Research 21(1) 1-31",
     "Persuasion knowledge", None, None, None, None, "LR14 Friestad & Wright 1994",
     "Forward citation pass – 30+ years of persuasion knowledge research. Filter to health/risk "
     "communication contexts for relevance to RQ1-RQ2."),
    ("Jan 2027", "Web of Science",
     "Cites: Slovic (1987) Science 236(4799) 280-285",
     "Risk perception", None, None, None, None, "LR04 Slovic 1987",
     "Forward citation pass – 30+ years of risk perception building on Slovic 1987."),
]

for row_idx, entry in enumerate(KW_SEARCHES, 49):
    row_fill = fill(ROW_EVEN) if row_idx % 2 == 0 else fill(ROW_ODD)
    for col_idx, val in enumerate(entry, 1):
        c = ws_lr.cell(row=row_idx, column=col_idx, value=val)
        if col_idx in (5, 6, 7, 8):  # numeric entry columns – editable
            c.fill = fill(EDIT_BG)
        else:
            c.fill = row_fill
        c.font = cell_font(size=9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = thin_border()
    ws_lr.row_dimensions[row_idx].height = 35

# ── Data validation: Status column K ─────────────────────────────────────────
dv_status = DataValidation(
    type="list",
    formula1='"To read,Reading,Read,Cited"',
    allow_blank=True,
    showDropDown=False
)
dv_status.sqref = f"K5:K{lr_data_end}"
dv_status.error      = "Choose: To read, Reading, Read, or Cited"
dv_status.errorTitle = "Invalid status"
dv_status.prompt     = "Track your reading status"
dv_status.promptTitle = "Status"
ws_lr.add_data_validation(dv_status)

# ── Conditional formatting on Status ─────────────────────────────────────────
green_fill = PatternFill("solid", fgColor="D1FAE5")
amber_fill = PatternFill("solid", fgColor="FEF3C7")
ws_lr.conditional_formatting.add(
    f"K5:K{lr_data_end}",
    CellIsRule(operator="equal", formula=['"Read"'], fill=green_fill)
)
ws_lr.conditional_formatting.add(
    f"K5:K{lr_data_end}",
    CellIsRule(operator="equal", formula=['"Cited"'], fill=green_fill)
)
ws_lr.conditional_formatting.add(
    f"K5:K{lr_data_end}",
    CellIsRule(operator="equal", formula=['"Reading"'], fill=amber_fill)
)

print("✓ Literature Review tab built: 23 papers pre-populated (LR01–LR23), "
      "16 blank rows (LR24–LR39), 20 keyword searches in Section 2")


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
wb.save(OUTPUT)
print(f"\n✓ Saved → {OUTPUT}")
