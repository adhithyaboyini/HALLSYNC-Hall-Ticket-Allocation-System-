# 🏫 HallSync — Examination Hall Ticket Allocation System

> A Python-based system that collects student data from multiple schools, validates it, allocates students to exam centres with smart seating rules, and generates professional hall ticket PDFs — fully automated.

---

## 📋 Table of Contents

- [What is HallSync?](#what-is-hallsync)
- [Features](#features)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [How to Use](#how-to-use)
- [Pipeline Overview](#pipeline-overview)
- [Allocation Rules](#allocation-rules)
- [Hall Ticket ID Format](#hall-ticket-id-format)
- [Output Files](#output-files)
- [Requirements](#requirements)

---

## What is HallSync?

HallSync automates the entire examination hall ticket process for schools and exam boards:

1. Each school fills a standard Excel template with student data
2. Admin drops all files into one folder
3. One command runs the full pipeline — aggregation → validation → allocation → PDF generation

---

## Features

- ✅ Accepts student data from **multiple schools** — single-file or multi-sheet Excel formats
- ✅ **11-point validation layer** — catches missing fields, duplicates, wrong formats before allocation
- ✅ **Smart allocation engine** — 4 rules enforced simultaneously
- ✅ **Auto hall calculation** — 30 seats per hall by default
- ✅ **4 hall tickets per A4 page** — one PDF per exam centre
- ✅ **8-digit unique Hall Ticket ID** — `YYDDCCSS` format
- ✅ **QR code** on every ticket
- ✅ **Submission status report** — tracks which schools submitted, which had errors
- ✅ Full **User Manual** included (Word document)

---

## Project Structure

```
HallSync/
│
├── main.py                          ← Run this — executes full pipeline
│
├── scripts/
│   ├── step0_aggregator.py          ← Merge all school Excel files
│   ├── step1_create_student_template.py  ← Generate student input template
│   ├── step2_create_school_template.py   ← Generate school info template
│   ├── step3_validation.py          ← Validate all data
│   ├── step4_allocation.py          ← Allocate students to centres & seats
│   └── step5_hall_tickets.py        ← Generate hall ticket PDFs
│
├── data/
│   ├── inbox/                       ← DROP school Excel files here
│   ├── Student_Info_Template.xlsm   ← Template given to each school
│   └── School_Info_Template.xlsx    ← Template filled by admin
│
├── output/
│   ├── allocation_result.xlsx       ← Generated after Step 4
│   └── hall_tickets/                ← PDFs generated after Step 5
│
├── vba/
│   ├── ThisWorkbook.bas             ← VBA code for subject multi-select
│   ├── Module1.bas                  ← VBA module
│   └── HOW_TO_INSTALL.txt          ← VBA setup guide
│
└── docs/
    └── HallSync_User_Manual.docx    ← Full user guide
```

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/HallSync.git
cd HallSync
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. (One-time) Generate the Excel templates

```bash
python scripts/step1_create_student_template.py
python scripts/step2_create_school_template.py
```

---

## How to Use

### Step 1 — Each school fills the student template

- Open `data/Student_Info_Template.xlsm` in **Microsoft Excel**
- Enable macros when prompted
- Fill one row per student (see User Manual for field instructions)
- Click any **Subjects cell** → multi-select popup appears
- Save the file

### Step 2 — Admin fills the school info template

- Open `data/School_Info_Template.xlsx`
- Add one row per school
- Set `Can_Be_Centre` to `Yes` or `No`
- Enter `Capacity` for all centre schools

### Step 3 — Drop all school files into inbox

```
data/inbox/
    SCH101_AlphaSchool.xlsx
    SCH102_BetaSchool.xlsx
    SCH103_GammaSchool.xlsx
    SCH104_SCH105_MultiSchool.xlsx    ← multi-sheet files also supported
```

### Step 4 — Run the full pipeline

```bash
python main.py
```

That's it. The system will:
- Merge all school files into a master database
- Validate every field
- Allocate students to centres and seats
- Generate hall ticket PDFs in `output/hall_tickets/`

---

## Pipeline Overview

```
data/inbox/  (school Excel files)
       ↓
Step 0 — Aggregation      →  data/Student_Info_Master.xlsx
                          →  data/Submission_Status_Report.xlsx
       ↓
Step 3 — Validation       →  Pass ✅ or Fail ❌ with detailed error report
       ↓
Step 4 — Allocation       →  output/allocation_result.xlsx
       ↓
Step 5 — Hall Tickets     →  output/hall_tickets/<Centre_Code>_HallTickets.pdf
```

---

## Allocation Rules

The allocation engine enforces 4 rules simultaneously:

| Rule | Description |
|------|-------------|
| **R1** | A student is **never** allocated to their own school's centre |
| **R2** | Students are allocated to centres in the **same district first** |
| **R3** | No two consecutive seats from the **same school** (best effort interleaving) |
| **R4** | **Gender alternation** M→F→M→F seat by seat (best effort) |

---

## Hall Ticket ID Format

Each hall ticket has a unique 8-digit ID: **`YYDDCCSS`**

| Segment | Meaning | Example |
|---------|---------|---------|
| `YY` | Last 2 digits of year | `26` |
| `DD` | District sequence number | `01` |
| `CC` | Centre sequence number | `02` |
| `SS` | Seat number | `07` |

Example: `26010207` = Year 2026, District 01, Centre 02, Seat 07

---

## Output Files

| File | Description |
|------|-------------|
| `data/Student_Info_Master.xlsx` | Merged master student database |
| `data/Submission_Status_Report.xlsx` | Which schools submitted, errors, duplicates |
| `output/allocation_result.xlsx` | Every student with centre, hall, seat assigned |
| `output/hall_tickets/*.pdf` | One PDF per centre, 4 tickets per page |

---

## Requirements

- Python 3.10+
- Microsoft Excel (for `.xlsm` macro features)
- See `requirements.txt` for Python packages

---

## License

MIT License — free to use, modify, and distribute.

---

*Built with Python · reportlab · openpyxl · pandas*
