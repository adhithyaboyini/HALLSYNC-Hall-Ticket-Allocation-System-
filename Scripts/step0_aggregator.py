"""
step0_aggregator.py  —  HallSync Data Aggregation Layer

Scans a single inbox folder, reads all school Excel files,
merges them into one Master Database, and produces a
Submission Status Report.

Supported input formats:
  A) Single-sheet file  — Student_Info sheet  (one school per file)
  B) Multi-sheet file   — one sheet per school (many schools, one file)

Usage:
    1. Drop all school Excel files into:   data/inbox/
    2. Run:  python step0_aggregator.py
    3. Outputs:
         data/Student_Info_Master.xlsx      ← feeds into Step 3 validation
         data/Submission_Status_Report.xlsx ← admin review
"""

import pandas as pd
import os, sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════════════════════════════════════════
# Configuration
# ════════════════════════════════════════════════════════════════════

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR  = os.path.join(BASE_DIR, "data", "inbox")
MASTER_OUT = os.path.join(BASE_DIR, "data", "Student_Info_Master.xlsx")
REPORT_OUT = os.path.join(BASE_DIR, "data", "Submission_Status_Report.xlsx")

os.makedirs(INBOX_DIR, exist_ok=True)

# Columns that MUST exist in every sheet
REQUIRED_COLS = [
    "Student_ID", "Student_Name", "Father_Name", "Mother_Name",
    "Gender", "Date_of_Birth", "Language_1", "Language_2",
    "Medium_of_Instruction", "Subjects",
    "School_Name", "School_Code", "Area", "District_Code"
]
# Columns that are kept if present but not required
OPTIONAL_COLS = ["Total_Subjects"]

# Colour palette
NAVY  = "1A3C6E"; BLUE  = "2E6DA4"; GREEN = "0D5C2F"
LIGHT = "EEF2F9"; WHITE = "FFFFFF"; GOLD  = "FFC107"
RED   = "C0392B"; AMBER = "FFF3CD"; GREY  = "555555"
GBGD  = "E8F5E9"; RBGD  = "FDECEA"


# ════════════════════════════════════════════════════════════════════
# Helpers
# ════════════════════════════════════════════════════════════════════

def tb(color="BBBBBB"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def hc(cell, text, bg=NAVY, fg=WHITE, sz=10, bold=True, wrap=False):
    cell.value     = text
    cell.font      = Font(bold=bold, size=sz, color=fg, name="Arial")
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                               wrap_text=wrap)
    cell.border    = tb()

def dc(cell, text="", bg=WHITE, bold=False, color="000000",
       center=False, sz=9):
    cell.value     = text
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.font      = Font(name="Arial", size=sz, bold=bold, color=color)
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center")
    cell.border    = tb()

def is_empty(value) -> bool:
    if value is None:
        return True
    return str(value).strip().lower() in ("", "nan", "none")


# ════════════════════════════════════════════════════════════════════
# File reader — handles both single-sheet and multi-sheet formats
# ════════════════════════════════════════════════════════════════════

class SchoolFileResult:
    """Holds the result of reading one school Excel file."""
    def __init__(self, filepath: str):
        self.filepath    = filepath
        self.filename    = os.path.basename(filepath)
        self.sheets      = []    # list of SheetResult
        self.file_error  = None  # set if file could not be opened at all

    @property
    def total_students(self):
        return sum(s.student_count for s in self.sheets if s.ok)

    @property
    def ok(self):
        return self.file_error is None and any(s.ok for s in self.sheets)


class SheetResult:
    """Holds the result of reading one sheet within a file."""
    def __init__(self, sheet_name: str):
        self.sheet_name    = sheet_name
        self.df            = None   # cleaned DataFrame if successful
        self.student_count = 0
        self.missing_cols  = []
        self.extra_cols    = []
        self.error         = None   # set if sheet could not be read
        self.skipped       = False  # True if intentionally skipped

    @property
    def ok(self):
        return (not self.skipped and self.error is None
                and len(self.missing_cols) == 0
                and self.student_count > 0)


def _clean_sheet_df(df: pd.DataFrame) -> pd.DataFrame:
    """Strip headers, drop empty rows, drop example row."""
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    if "Student_ID" in df.columns:
        df = df[~df["Student_ID"].astype(str).str.strip().str.lower()
                  .isin(["s001", "nan", "none", "student_id", ""])]
    return df.reset_index(drop=True)


def _detect_header_row(df_raw: pd.DataFrame) -> int:
    """
    Find which row contains the actual column headers.
    Looks for a row that contains 'Student_ID' or 'Student_Name'.
    Returns the 0-based row index to use as header, default 4.
    """
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values]
        if "Student_ID" in vals or "Student_Name" in vals:
            return i
    return 4   # default: HallSync template header row


def read_sheet(filepath: str, sheet_name: str) -> SheetResult:
    """Read one sheet and return a SheetResult."""
    result = SheetResult(sheet_name)

    # Skip obviously decorative sheet names
    skip_names = {"instructions", "subjectlist", "readme",
                  "sheet1", "sheet2", "sheet3"}
    if sheet_name.strip().lower() in skip_names:
        result.skipped = True
        return result

    try:
        # Read raw first to detect header row
        df_raw = pd.read_excel(filepath, sheet_name=sheet_name,
                               header=None, dtype=str)
        header_row = _detect_header_row(df_raw)
        df = pd.read_excel(filepath, sheet_name=sheet_name,
                           header=header_row, dtype=str)
        df = _clean_sheet_df(df)

        if df.empty:
            result.skipped = True
            result.error   = "Sheet is empty after cleaning"
            return result

        # Check columns
        present     = set(df.columns.tolist())
        missing     = [c for c in REQUIRED_COLS if c not in present]
        extra       = [c for c in present if c not in REQUIRED_COLS
                       and c not in OPTIONAL_COLS]

        result.missing_cols = missing
        result.extra_cols   = extra

        if not missing:
            # Keep required + any optional columns that are present
            keep = [c for c in REQUIRED_COLS if c in df.columns]
            for oc in OPTIONAL_COLS:
                if oc in df.columns:
                    keep.append(oc)
            df             = df[keep]
            result.df      = df
            result.student_count = len(df)
        else:
            result.error = f"Missing columns: {', '.join(missing)}"

    except Exception as e:
        result.error = str(e)

    return result


def read_school_file(filepath: str) -> SchoolFileResult:
    """
    Read one school Excel file.
    Tries all sheets and returns results for each.
    """
    file_result = SchoolFileResult(filepath)

    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names
    except Exception as e:
        file_result.file_error = f"Cannot open file: {e}"
        return file_result

    for sname in sheet_names:
        sr = read_sheet(filepath, sname)
        file_result.sheets.append(sr)

    return file_result


# ════════════════════════════════════════════════════════════════════
# Aggregator
# ════════════════════════════════════════════════════════════════════

def aggregate(inbox_dir: str) -> tuple[pd.DataFrame, list[SchoolFileResult], dict]:
    """
    Scan inbox_dir, read all Excel files, merge into one DataFrame.

    Returns:
        master_df     : merged student DataFrame
        file_results  : list of SchoolFileResult (one per file)
        stats         : summary statistics dict
    """
    print(f"\n  Scanning inbox: {inbox_dir}")

    # Find all Excel files
    all_files = []
    for fname in sorted(os.listdir(inbox_dir)):
        if fname.startswith("~$"):          # skip Excel lock files
            continue
        if fname.lower().endswith((".xlsx", ".xlsm", ".xls")):
            all_files.append(os.path.join(inbox_dir, fname))

    print(f"  Found {len(all_files)} file(s)\n")

    file_results   = []
    all_dfs        = []
    seen_ids       = {}   # student_id → (filename, sheet)
    duplicate_ids  = []

    for fpath in all_files:
        fname = os.path.basename(fpath)
        print(f"  [→] Reading: {fname}")
        fr = read_school_file(fpath)
        file_results.append(fr)

        if fr.file_error:
            print(f"      ❌ File error: {fr.file_error}")
            continue

        for sr in fr.sheets:
            if sr.skipped:
                continue
            if sr.error:
                print(f"      ⚠️  Sheet '{sr.sheet_name}': {sr.error}")
                continue
            if not sr.ok:
                continue

            # Check for cross-file duplicate Student_IDs
            for sid in sr.df["Student_ID"].astype(str).str.strip():
                sid_upper = sid.upper()
                if sid_upper in seen_ids:
                    duplicate_ids.append({
                        "Student_ID": sid,
                        "File_1":  seen_ids[sid_upper][0],
                        "Sheet_1": seen_ids[sid_upper][1],
                        "File_2":  fname,
                        "Sheet_2": sr.sheet_name,
                    })
                else:
                    seen_ids[sid_upper] = (fname, sr.sheet_name)

            # Tag source file and sheet
            df_tagged = sr.df.copy()
            df_tagged["_source_file"]  = fname
            df_tagged["_source_sheet"] = sr.sheet_name
            all_dfs.append(df_tagged)

            print(f"      ✅ Sheet '{sr.sheet_name}': "
                  f"{sr.student_count} students")

    # Merge all DataFrames
    if all_dfs:
        master_df = pd.concat(all_dfs, ignore_index=True)
        # Drop source tags before saving (keep clean for validator)
        master_clean = master_df.drop(
            columns=["_source_file", "_source_sheet"], errors="ignore"
        )
    else:
        master_df    = pd.DataFrame(columns=REQUIRED_COLS)
        master_clean = master_df.copy()

    stats = {
        "total_files":      len(all_files),
        "ok_files":         sum(1 for fr in file_results if fr.ok),
        "error_files":      sum(1 for fr in file_results if not fr.ok),
        "total_students":   len(master_df),
        "duplicate_ids":    duplicate_ids,
        "timestamp":        datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    }

    return master_clean, master_df, file_results, stats


# ════════════════════════════════════════════════════════════════════
# Save Master Database
# ════════════════════════════════════════════════════════════════════

def save_master(master_df: pd.DataFrame, out_path: str):
    """
    Save the merged student data as a styled Excel file
    that is directly compatible with step3_validation.py
    (same sheet name and header row structure as the template).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Student_Info"

    # Row 1 — Title
    ws.merge_cells("A1:N1")
    hc(ws["A1"],
       f"🏫  HallSync — Master Student Database  "
       f"(Generated: {datetime.now().strftime('%d-%m-%Y %H:%M')})",
       bg=NAVY, sz=12)
    ws.row_dimensions[1].height = 28

    # Row 2 — Subtitle
    ws.merge_cells("A2:N2")
    c2 = ws["A2"]
    c2.value = (f"Aggregated from inbox folder  •  "
                f"{len(master_df)} total students  •  "
                f"Feed this file into step3_validation.py")
    c2.font      = Font(italic=True, size=9, color=GREY, name="Arial")
    c2.fill      = PatternFill("solid", start_color=AMBER)
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # Row 3 — Gold line
    for col in range(1, 15):
        ws.cell(3, col).fill = PatternFill("solid", start_color=GOLD)
    ws.row_dimensions[3].height = 4

    # Row 4 — Section labels
    ws.merge_cells("A4:F4")
    hc(ws["A4"], "👤  Student Details", bg=NAVY)
    ws.merge_cells("G4:J4")
    hc(ws["G4"], "📚  Academic Details", bg=BLUE)
    ws.merge_cells("K4:N4")
    hc(ws["K4"], "🏫  School Details", bg=GREEN)
    ws.row_dimensions[4].height = 20

    # Row 5 — Headers
    COLS = [
        ("Student_ID",            14, NAVY),
        ("Student_Name",          24, NAVY),
        ("Father_Name",           22, NAVY),
        ("Mother_Name",           22, NAVY),
        ("Gender",                12, NAVY),
        ("Date_of_Birth",         16, NAVY),
        ("Language_1",            16, BLUE),
        ("Language_2",            16, BLUE),
        ("Medium_of_Instruction", 22, BLUE),
        ("Subjects",              32, BLUE),
        ("School_Name",           30, GREEN),
        ("School_Code",           14, GREEN),
        ("Area",                  20, GREEN),
        ("District_Code",         16, GREEN),
    ]
    for i, (name, width, bg) in enumerate(COLS, 1):
        hc(ws.cell(5, i), name, bg=bg, sz=9, wrap=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[5].height = 28

    # Data rows (starting row 6 — no example row in master)
    COL_NAMES = [c[0] for c in COLS]
    for r_idx, (_, row) in enumerate(master_df.iterrows(), 6):
        bg = LIGHT if r_idx % 2 == 0 else WHITE
        for c_idx, col in enumerate(COL_NAMES, 1):
            val  = str(row.get(col, "")).strip()
            val  = "" if val.lower() in ("nan", "none") else val
            cell = ws.cell(r_idx, c_idx, val)
            cell.fill      = PatternFill("solid", start_color=bg)
            cell.font      = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center",
                                       horizontal="center"
                                       if c_idx in (1,5,6,12) else "left")
            cell.border    = tb()
        ws.row_dimensions[r_idx].height = 17

    ws.freeze_panes = "A6"
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    print(f"\n  [✓] Master database saved → {out_path}")


# ════════════════════════════════════════════════════════════════════
# Save Submission Status Report
# ════════════════════════════════════════════════════════════════════

def save_report(file_results: list[SchoolFileResult],
                master_df: pd.DataFrame,
                stats: dict,
                out_path: str):

    wb = Workbook()

    # ── Sheet 1: Dashboard ────────────────────────────────────────────────────
    wd = wb.active
    wd.title = "Dashboard"
    wd.sheet_view.showGridLines = False

    wd.merge_cells("A1:F1")
    hc(wd["A1"], "📊  HallSync — Submission Status Dashboard", bg=NAVY, sz=13)
    wd.row_dimensions[1].height = 30

    # Timestamp
    wd.merge_cells("A2:F2")
    t = wd["A2"]
    t.value     = f"Generated: {stats['timestamp']}"
    t.font      = Font(italic=True, size=9, color=GREY, name="Arial")
    t.alignment = Alignment(horizontal="center")
    wd.row_dimensions[2].height = 16

    for col in range(1, 7):
        wd.cell(3, col).fill = PatternFill("solid", start_color=GOLD)
    wd.row_dimensions[3].height = 4

    # KPIs
    kpis = [
        ("Total Files",     stats["total_files"],    NAVY),
        ("Processed OK",    stats["ok_files"],       GREEN),
        ("Errors/Skipped",  stats["error_files"],    RED if stats["error_files"] > 0 else GREEN),
        ("Total Students",  stats["total_students"], BLUE),
        ("Duplicate IDs",   len(stats["duplicate_ids"]),
                            RED if stats["duplicate_ids"] else GREEN),
    ]
    for col_offset, (label, value, color) in enumerate(kpis, 1):
        lc = wd.cell(5, col_offset, label)
        lc.font      = Font(bold=True, size=8, color=GREY, name="Arial")
        lc.alignment = Alignment(horizontal="center")
        vc = wd.cell(6, col_offset, value)
        vc.font      = Font(bold=True, size=20, color=color, name="Arial")
        vc.fill      = PatternFill("solid", start_color=LIGHT)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = tb()
        wd.column_dimensions[get_column_letter(col_offset)].width = 18
    wd.row_dimensions[5].height = 16
    wd.row_dimensions[6].height = 36

    # File summary table
    wd.cell(8, 1, "File Summary").font = Font(bold=True, size=11,
                                              color=NAVY, name="Arial")
    wd.row_dimensions[8].height = 18

    hdrs = ["File Name", "Sheets Read", "Students", "Status", "Issues"]
    col_widths = [36, 14, 12, 16, 40]
    for ci, (h, w) in enumerate(zip(hdrs, col_widths), 1):
        hc(wd.cell(9, ci), h, bg=BLUE, sz=9)
        wd.column_dimensions[get_column_letter(ci)].width = w
    wd.row_dimensions[9].height = 20

    for ri, fr in enumerate(file_results, 10):
        bg = LIGHT if ri % 2 == 0 else WHITE

        ok_sheets = sum(1 for s in fr.sheets if s.ok)
        issues    = []

        if fr.file_error:
            status = "❌ File Error"
            issues.append(fr.file_error)
        elif fr.total_students == 0:
            status = "⚠️  Empty"
            issues.append("No valid student records found")
        else:
            status = "✅ OK"

        for s in fr.sheets:
            if s.missing_cols:
                issues.append(f"Sheet '{s.sheet_name}': "
                               f"missing {', '.join(s.missing_cols)}")
            if s.error and not s.skipped:
                issues.append(f"Sheet '{s.sheet_name}': {s.error}")

        status_bg = (GBGD if "✅" in status
                     else RBGD if "❌" in status else AMBER)

        dc(wd.cell(ri, 1), fr.filename,          bg=bg,       bold=True)
        dc(wd.cell(ri, 2), ok_sheets,             bg=bg,       center=True)
        dc(wd.cell(ri, 3), fr.total_students,     bg=bg,       center=True,
           bold=True, color=BLUE)
        dc(wd.cell(ri, 4), status,                bg=status_bg, center=True)
        dc(wd.cell(ri, 5), "  |  ".join(issues) if issues else "—",
           bg=bg, color=RED if issues else GREY)
        wd.row_dimensions[ri].height = 18

    # ── Sheet 2: All Students ─────────────────────────────────────────────────
    ws = wb.create_sheet("All Students")
    if not master_df.empty:
        hdrs2 = list(master_df.columns)
        for ci, h in enumerate(hdrs2, 1):
            hc(ws.cell(1, ci), h, bg=NAVY, sz=9)
            ws.column_dimensions[get_column_letter(ci)].width = 18
        ws.row_dimensions[1].height = 20

        for ri, (_, row) in enumerate(master_df.iterrows(), 2):
            bg = LIGHT if ri % 2 == 0 else WHITE
            for ci, col in enumerate(hdrs2, 1):
                val = str(row.get(col, "")).strip()
                val = "" if val.lower() in ("nan","none") else val
                dc(ws.cell(ri, ci), val, bg=bg)
            ws.row_dimensions[ri].height = 16
        ws.freeze_panes = "A2"

    # ── Sheet 3: Duplicate IDs ────────────────────────────────────────────────
    dup_sheet = wb.create_sheet("⚠ Duplicate IDs")
    if stats["duplicate_ids"]:
        dup_hdrs = ["Student_ID", "File_1", "Sheet_1", "File_2", "Sheet_2"]
        for ci, h in enumerate(dup_hdrs, 1):
            hc(dup_sheet.cell(1, ci), h, bg=RED, sz=9)
            dup_sheet.column_dimensions[get_column_letter(ci)].width = 28
        for ri, dup in enumerate(stats["duplicate_ids"], 2):
            bg = RBGD if ri % 2 == 0 else WHITE
            for ci, key in enumerate(dup_hdrs, 1):
                dc(dup_sheet.cell(ri, ci), dup.get(key, ""), bg=bg,
                   color=RED, bold=(ci==1))
            dup_sheet.row_dimensions[ri].height = 16
    else:
        dup_sheet.merge_cells("A1:E1")
        hc(dup_sheet["A1"],
           "✅  No duplicate Student_IDs found across all files.",
           bg=GREEN, sz=11)
        dup_sheet.row_dimensions[1].height = 28

    # ── Sheet 4: School Breakdown ─────────────────────────────────────────────
    sb = wb.create_sheet("School Breakdown")
    if not master_df.empty and "School_Code" in master_df.columns:
        breakdown = (master_df.groupby(
                        ["School_Code","School_Name","District_Code"],
                        dropna=False)
                     .agg(Students=("Student_ID","count"))
                     .reset_index()
                     .sort_values("School_Code"))

        bhdrs = ["School_Code","School_Name","District_Code","Students"]
        for ci, h in enumerate(bhdrs, 1):
            hc(sb.cell(1, ci), h, bg=BLUE, sz=9)
            sb.column_dimensions[get_column_letter(ci)].width = [14,36,16,12][ci-1]
        sb.row_dimensions[1].height = 20

        for ri, (_, row) in enumerate(breakdown.iterrows(), 2):
            bg = LIGHT if ri % 2 == 0 else WHITE
            dc(sb.cell(ri,1), row["School_Code"],   bg=bg, bold=True, color=NAVY)
            dc(sb.cell(ri,2), row["School_Name"],   bg=bg)
            dc(sb.cell(ri,3), row["District_Code"], bg=bg)
            dc(sb.cell(ri,4), int(row["Students"]), bg=bg,
               bold=True, color=BLUE, center=True)
            sb.row_dimensions[ri].height = 16

        # Total row
        tr = len(breakdown) + 2
        hc(sb.cell(tr, 1), "TOTAL", bg=NAVY, sz=9)
        hc(sb.cell(tr, 2), "", bg=NAVY)
        hc(sb.cell(tr, 3), "", bg=NAVY)
        hc(sb.cell(tr, 4), int(master_df["Student_ID"].count()), bg=NAVY, sz=10)
        sb.row_dimensions[tr].height = 20

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    print(f"  [✓] Status report saved  → {out_path}")


# ════════════════════════════════════════════════════════════════════
# Print terminal summary
# ════════════════════════════════════════════════════════════════════

def print_summary(file_results, stats):
    print(f"\n{'═'*65}")
    print(f"  HallSync — Aggregation Summary")
    print(f"{'═'*65}")
    print(f"  Files found       : {stats['total_files']}")
    print(f"  Files processed   : {stats['ok_files']}")
    print(f"  Files with errors : {stats['error_files']}")
    print(f"  Total students    : {stats['total_students']}")
    print(f"  Duplicate IDs     : {len(stats['duplicate_ids'])}")

    if stats["duplicate_ids"]:
        print(f"\n  ⚠️  Duplicate Student_IDs (must fix before validation):")
        for d in stats["duplicate_ids"]:
            print(f"     {d['Student_ID']}  —  "
                  f"{d['File_1']} / {d['Sheet_1']}  vs  "
                  f"{d['File_2']} / {d['Sheet_2']}")

    print(f"\n  Per-file status:")
    print(f"  {'File':<38} {'Students':>9}  Status")
    print(f"  {'─'*60}")
    for fr in file_results:
        icon = "✅" if fr.ok else ("❌" if fr.file_error else "⚠️ ")
        print(f"  {icon}  {fr.filename:<36} {fr.total_students:>9}")

    print(f"{'═'*65}")

    if stats["duplicate_ids"]:
        print(f"\n  ❌  Fix duplicate Student_IDs before running validation.")
    elif stats["total_students"] == 0:
        print(f"\n  ⚠️  No students found. Check the inbox folder.")
    else:
        print(f"\n  ✅  Master database ready.")
        print(f"      Next step: python step3_validation.py")
    print()


# ════════════════════════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════════════════════════

def run():
    print(f"\n{'═'*65}")
    print(f"  HallSync — Step 0: Data Aggregation Layer")
    print(f"{'═'*65}")
    print(f"\n  Inbox folder : {INBOX_DIR}")
    print(f"  Drop all school Excel files into the inbox folder,")
    print(f"  then re-run this script.\n")

    if not os.listdir(INBOX_DIR):
        print(f"  ⚠️  Inbox is empty. No files to process.")
        print(f"      Add school Excel files to: {INBOX_DIR}\n")
        return

    # Aggregate
    master_clean, master_tagged, file_results, stats = aggregate(INBOX_DIR)

    # Save master database
    if stats["total_students"] > 0:
        save_master(master_clean, MASTER_OUT)
    else:
        print("\n  ⚠️  No students to save.")

    # Save report
    save_report(file_results, master_tagged, stats, REPORT_OUT)

    # Print terminal summary
    print_summary(file_results, stats)


if __name__ == "__main__":
    run()
