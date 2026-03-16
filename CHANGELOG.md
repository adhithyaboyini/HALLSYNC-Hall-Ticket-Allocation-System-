# Changelog

## v1.0.0 — March 2026

### Initial Release

**Data Collection**
- Student_Info_Template.xlsm with 15 fields across 3 sections
- Multi-select subject dropdown via VBA macro
- Auto-calculated Total_Subjects formula
- School_Info_Template.xlsx for admin use

**Aggregation Layer (Step 0)**
- Scans inbox folder for all Excel files
- Supports single-school files and multi-sheet files
- Cross-file duplicate Student_ID detection
- Submission Status Report with 4-sheet Excel output

**Validation Layer (Step 3)**
- 11-point validation covering all fields
- Logical date validation (DD-MM-YYYY + real date check)
- Case-insensitive School_Code matching
- Global capacity vs student count check
- Exits with code 0 (pass) or 1 (fail) for pipeline chaining

**Allocation Engine (Step 4)**
- R1: Students never allocated to own school
- R2: Same-district centre preference
- R3: Best-effort school interleaving (no consecutive same-school seats)
- R4: Best-effort gender alternation M→F→M→F
- Auto hall generation at 30 seats per hall
- Post-allocation rule verification report

**Hall Ticket Generator (Step 5)**
- 4 tickets per A4 page (2×2 grid)
- One PDF per exam centre
- 8-digit unique Hall Ticket ID (YYDDCCSS format)
- QR code on every ticket
- Photo placeholder
- Signature lines

**Documentation**
- Full User Manual (Word document)
- SETUP.md first-time installation guide
- README.md with full pipeline documentation
