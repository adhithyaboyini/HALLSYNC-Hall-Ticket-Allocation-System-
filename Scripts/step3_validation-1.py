"""
step3_validation.py  —  HallSync Step 3: Validation Layer

Bug fixes applied:
  BUG 1  - excel_row off-by-one for students (example row dropped → idx+7 not idx+6)
  BUG 2  - nan check missed 'NaN','NAN',None → unified is_empty() helper
  BUG 3  - Can_Be_Centre: .capitalize() on 'NaN' gave 'Nan', bypassing empty check
  BUG 4  - Gender 'NaN' gave WARNING instead of ERROR (missing)
  BUG 5  - Fractional capacity '120.5' silently truncated with no warning
  BUG 6  - None passed all nan checks (str(None)='None' not 'nan')
  BUG 7  - Loader filter case-sensitive for example row ('S001' vs 's001')
  BUG 8  - School_Code case mismatch: student 'sch101' not matched to school 'SCH101'
  BUG 9  - Duplicate School_Code check case-insensitive: 'SCH101'+'sch101' not caught
  BUG 10 - DOB regex allows logically invalid dates: month 13, day 32, 00-00-0000
  BUG 11 - Empty student file passes silently: total_students=0 skips capacity check
"""

import pandas as pd
import re
import os
import sys
from dataclasses import dataclass, field
from typing import List
from datetime import datetime


# ════════════════════════════════════════════════════════════════════
# Data classes
# ════════════════════════════════════════════════════════════════════

@dataclass
class Issue:
    row:     int
    field:   str
    message: str
    level:   str

    def __str__(self):
        icon = "❌" if self.level == "ERROR" else "⚠️ "
        return f"  {icon}  Row {self.row:<5} | {self.field:<22} | {self.message}"


@dataclass
class ValidationReport:
    errors:   List[Issue] = field(default_factory=list)
    warnings: List[Issue] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return len(self.errors) == 0

    def add_error(self, row, field_name, msg):
        self.errors.append(Issue(row, field_name, msg, "ERROR"))

    def add_warning(self, row, field_name, msg):
        self.warnings.append(Issue(row, field_name, msg, "WARNING"))


# ════════════════════════════════════════════════════════════════════
# Loaders
# ════════════════════════════════════════════════════════════════════

def load_student_info(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Student_Info", header=4, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    # BUG FIX 7: case-insensitive filter for example row and empty values
    df = df[~df["Student_ID"].astype(str).str.strip().str.lower().isin(
        ["s001", "nan", "none", ""]
    )]
    return df.reset_index(drop=True)


def load_school_info(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="School_Info", header=4, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    df = df[~df["School_Code"].astype(str).str.strip().str.lower().isin(
        ["nan", "none", "", "totals"]
    )]
    df = df[~df["School_Code"].astype(str).str.startswith("=")]
    return df.reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════
# Constants
# ════════════════════════════════════════════════════════════════════

REQUIRED_STUDENT_COLS = [
    "Student_ID", "Student_Name", "Father_Name", "Mother_Name",
    "Gender", "Date_of_Birth", "Language_1", "Language_2",
    "Medium_of_Instruction", "Subjects", "School_Name",
    "School_Code", "Area", "District_Code"
]
REQUIRED_SCHOOL_COLS = [
    "School_Code", "School_Name", "Area",
    "District_Code", "Can_Be_Centre", "Capacity"
]
VALID_GENDERS = {"male", "female", "other"}
VALID_MEDIUMS = {"tamil medium", "english medium", "telugu medium",
                 "kannada medium", "hindi medium", "malayalam medium"}
DOB_PATTERN   = re.compile(r"^\d{2}-\d{2}-\d{4}$")


# ════════════════════════════════════════════════════════════════════
# Helper: unified empty check  (BUG FIX 2, 6)
# ════════════════════════════════════════════════════════════════════

def is_empty(value) -> bool:
    """Catches None, '', ' ', 'nan', 'NaN', 'NAN', 'none', 'None', 'NONE'."""
    if value is None:
        return True
    return str(value).strip().lower() in ("", "nan", "none")


# ════════════════════════════════════════════════════════════════════
# Helper: logical date validation  (BUG FIX 10)
# ════════════════════════════════════════════════════════════════════

def is_valid_dob(dob: str) -> tuple[bool, str]:
    """
    Returns (is_valid, reason).
    Checks format AND logical validity (no month 13, day 32, all-zeros).
    BUG FIX 10: regex alone allowed '32-13-0000' to pass.
    """
    if not DOB_PATTERN.match(dob):
        return False, f"Invalid format '{dob}' — expected DD-MM-YYYY"
    try:
        parsed = datetime.strptime(dob, "%d-%m-%Y")
        # Reject obviously wrong years
        if parsed.year < 1900 or parsed.year > 2100:
            return False, f"Year {parsed.year} is out of range"
        return True, ""
    except ValueError:
        return False, f"'{dob}' is not a real date (e.g. day 32 or month 13)"


# ════════════════════════════════════════════════════════════════════
# Student Validation
# ════════════════════════════════════════════════════════════════════

def validate_students(df: pd.DataFrame, school_codes: set) -> ValidationReport:
    report = ValidationReport()

    missing_cols = [c for c in REQUIRED_STUDENT_COLS if c not in df.columns]
    if missing_cols:
        for col in missing_cols:
            report.add_error(0, col,
                             f"Required column '{col}' is missing from Student_Info")
        return report

    # BUG FIX 8: normalise school_codes to uppercase for case-insensitive matching
    school_codes_upper = {s.upper() for s in school_codes}

    # BUG FIX 9: track seen IDs normalised to uppercase for case-insensitive dup check
    seen_ids = {}

    for idx, row in df.iterrows():
        # BUG FIX 1: example row is dropped before reset, so idx=0 = Excel row 7
        excel_row = idx + 7

        sid = str(row.get("Student_ID", "")).strip()

        # ── Student_ID ────────────────────────────────────────────────────────
        if is_empty(sid):
            report.add_error(excel_row, "Student_ID", "Student_ID is missing")
            continue

        sid_upper = sid.upper()
        if sid_upper in seen_ids:
            report.add_error(excel_row, "Student_ID",
                             f"Duplicate — '{sid}' already seen at Row {seen_ids[sid_upper]}")
        else:
            seen_ids[sid_upper] = excel_row

        # ── Student_Name ──────────────────────────────────────────────────────
        name = str(row.get("Student_Name", "")).strip()
        if is_empty(name):
            report.add_error(excel_row, "Student_Name",
                             f"Student_Name is missing for {sid}")
        elif len(name) < 2:
            report.add_warning(excel_row, "Student_Name",
                               f"Name '{name}' seems too short for {sid}")

        # ── Father_Name ───────────────────────────────────────────────────────
        if is_empty(row.get("Father_Name")):
            report.add_error(excel_row, "Father_Name",
                             f"Father_Name is missing for {sid}")

        # ── Mother_Name ───────────────────────────────────────────────────────
        if is_empty(row.get("Mother_Name")):
            report.add_error(excel_row, "Mother_Name",
                             f"Mother_Name is missing for {sid}")

        # ── Gender (BUG FIX 4) ────────────────────────────────────────────────
        gender = str(row.get("Gender", "")).strip()
        if is_empty(gender):
            report.add_error(excel_row, "Gender",
                             f"Gender is missing for {sid}")
        elif gender.lower() not in VALID_GENDERS:
            report.add_warning(excel_row, "Gender",
                               f"Unexpected value '{gender}' — expected Male / Female / Other")

        # ── Date of Birth (BUG FIX 10) ────────────────────────────────────────
        dob = str(row.get("Date_of_Birth", "")).strip()
        if is_empty(dob):
            report.add_error(excel_row, "Date_of_Birth",
                             f"Date_of_Birth is missing for {sid}")
        else:
            valid, reason = is_valid_dob(dob)
            if not valid:
                report.add_error(excel_row, "Date_of_Birth", f"{reason} for {sid}")

        # ── Language 1 ────────────────────────────────────────────────────────
        if is_empty(row.get("Language_1")):
            report.add_error(excel_row, "Language_1",
                             f"Language_1 is missing for {sid}")

        # ── Language 2 ────────────────────────────────────────────────────────
        if is_empty(row.get("Language_2")):
            report.add_error(excel_row, "Language_2",
                             f"Language_2 is missing for {sid}")

        # ── Medium of Instruction ─────────────────────────────────────────────
        medium = str(row.get("Medium_of_Instruction", "")).strip()
        if is_empty(medium):
            report.add_error(excel_row, "Medium_of_Instruction",
                             f"Medium_of_Instruction is missing for {sid}")
        elif medium.lower() not in VALID_MEDIUMS:
            report.add_warning(excel_row, "Medium_of_Instruction",
                               f"Unrecognised medium '{medium}' for {sid}")

        # ── Subjects ──────────────────────────────────────────────────────────
        subjects = str(row.get("Subjects", "")).strip()
        if is_empty(subjects):
            report.add_error(excel_row, "Subjects",
                             f"Subjects is missing for {sid}")
        else:
            subj_list = [s.strip() for s in subjects.split(",") if s.strip()]
            if len(subj_list) == 0:
                report.add_error(excel_row, "Subjects",
                                 f"Subjects cell is empty for {sid}")
            elif len(subj_list) < 2:
                report.add_warning(excel_row, "Subjects",
                                   f"Only 1 subject found for {sid} — is this correct?")

        # ── School_Code (BUG FIX 8) ───────────────────────────────────────────
        scode = str(row.get("School_Code", "")).strip()
        if is_empty(scode):
            report.add_error(excel_row, "School_Code",
                             f"School_Code is missing for {sid}")
        elif scode.upper() not in school_codes_upper:
            report.add_error(excel_row, "School_Code",
                             f"'{scode}' not found in School_Info — add it first")

        # ── School_Name ───────────────────────────────────────────────────────
        if is_empty(row.get("School_Name")):
            report.add_error(excel_row, "School_Name",
                             f"School_Name is missing for {sid}")

        # ── Area ──────────────────────────────────────────────────────────────
        if is_empty(row.get("Area")):
            report.add_warning(excel_row, "Area", f"Area is missing for {sid}")

        # ── District_Code ─────────────────────────────────────────────────────
        if is_empty(row.get("District_Code")):
            report.add_warning(excel_row, "District_Code",
                               f"District_Code is missing for {sid}")

    return report


# ════════════════════════════════════════════════════════════════════
# School Validation
# ════════════════════════════════════════════════════════════════════

def validate_schools(df: pd.DataFrame, total_students: int) -> ValidationReport:
    report = ValidationReport()

    missing_cols = [c for c in REQUIRED_SCHOOL_COLS if c not in df.columns]
    if missing_cols:
        for col in missing_cols:
            report.add_error(0, col,
                             f"Required column '{col}' is missing from School_Info")
        return report

    # BUG FIX 9: normalise to uppercase for case-insensitive duplicate detection
    seen_codes     = {}
    total_capacity = 0
    centre_count   = 0

    for idx, row in df.iterrows():
        # BUG FIX 7: school has no example row → idx=0 = Excel row 6
        excel_row = idx + 6

        scode = str(row.get("School_Code", "")).strip()

        # ── School_Code ───────────────────────────────────────────────────────
        if is_empty(scode):
            report.add_error(excel_row, "School_Code", "School_Code is missing")
            continue

        # BUG FIX 9: compare uppercased to catch 'SCH101' vs 'sch101' duplicates
        scode_upper = scode.upper()
        if scode_upper in seen_codes:
            report.add_error(excel_row, "School_Code",
                             f"Duplicate — '{scode}' already seen at Row {seen_codes[scode_upper]}")
        else:
            seen_codes[scode_upper] = excel_row

        # ── School_Name ───────────────────────────────────────────────────────
        if is_empty(row.get("School_Name")):
            report.add_error(excel_row, "School_Name",
                             f"School_Name is missing for {scode}")

        # ── Area ──────────────────────────────────────────────────────────────
        if is_empty(row.get("Area")):
            report.add_warning(excel_row, "Area", f"Area is missing for {scode}")

        # ── District_Code ─────────────────────────────────────────────────────
        if is_empty(row.get("District_Code")):
            report.add_warning(excel_row, "District_Code",
                               f"District_Code is missing for {scode}")

        # ── Can_Be_Centre (BUG FIX 3) ────────────────────────────────────────
        centre_raw = str(row.get("Can_Be_Centre", "")).strip()
        if is_empty(centre_raw):
            report.add_error(excel_row, "Can_Be_Centre",
                             f"Can_Be_Centre is missing for {scode} — must be Yes or No")
            centre_raw = ""
        else:
            centre = centre_raw.capitalize()
            if centre not in ("Yes", "No"):
                report.add_error(excel_row, "Can_Be_Centre",
                                 f"Invalid value '{centre_raw}' for {scode} — must be Yes or No")
                centre_raw = ""

        # ── Capacity (BUG FIX 5) ─────────────────────────────────────────────
        cap_raw = str(row.get("Capacity", "")).strip()
        if centre_raw.capitalize() == "Yes":
            if is_empty(cap_raw):
                report.add_error(excel_row, "Capacity",
                                 f"Capacity is missing for {scode} (Can_Be_Centre = Yes)")
            else:
                try:
                    cap_float = float(cap_raw)
                    # BUG FIX 5: warn on fractional capacity before truncating
                    if cap_float != int(cap_float):
                        report.add_warning(excel_row, "Capacity",
                                           f"Capacity '{cap_raw}' is fractional for {scode}"
                                           f" — rounded down to {int(cap_float)}")
                    cap = int(cap_float)
                    if cap <= 0:
                        report.add_error(excel_row, "Capacity",
                                         f"Capacity must be greater than 0 for centre {scode}")
                    else:
                        total_capacity += cap
                        centre_count   += 1
                except ValueError:
                    report.add_error(excel_row, "Capacity",
                                     f"Capacity '{cap_raw}' is not a valid number for {scode}")

    # ── Global checks ─────────────────────────────────────────────────────────
    if centre_count == 0:
        report.add_error(0, "Can_Be_Centre",
                         "No school has Can_Be_Centre = Yes — at least one centre is required")

    # BUG FIX 11: guard against total_students = 0 (empty file passed silently)
    if total_students == 0:
        report.add_error(0, "Student_Info",
                         "No student records found — Student_Info appears to be empty")
    elif centre_count > 0:
        if total_capacity < total_students:
            report.add_error(0, "Capacity",
                             f"Total capacity ({total_capacity}) is less than "
                             f"total students ({total_students}) — add more centres or increase capacity")
        elif total_capacity < total_students * 1.1:
            report.add_warning(0, "Capacity",
                               f"Total capacity ({total_capacity}) is less than 110% of "
                               f"students ({total_students}) — allocation may be tight")

    return report


# ════════════════════════════════════════════════════════════════════
# Report Printer
# ════════════════════════════════════════════════════════════════════

def print_report(title: str, report: ValidationReport):
    print(f"\n{'═'*65}")
    print(f"  {title}")
    print(f"{'═'*65}")
    if report.errors:
        print(f"\n  ❌  {len(report.errors)} ERROR(S) — must fix before proceeding\n")
        print(f"  {'':>4}  {'Row':<8}{'Field':<24}Issue")
        print(f"  {'─'*60}")
        for e in report.errors:
            print(str(e))
    else:
        print(f"\n  ✅  No errors found.")
    if report.warnings:
        print(f"\n  ⚠️   {len(report.warnings)} WARNING(S) — review recommended\n")
        for w in report.warnings:
            print(str(w))
    print(f"{'═'*65}")


# ════════════════════════════════════════════════════════════════════
# Main runner
# ════════════════════════════════════════════════════════════════════

def run_validation(student_path: str, school_path: str) -> bool:
    print(f"\n{'═'*65}")
    print(f"  HallSync — Step 3: Validation Layer")
    print(f"{'═'*65}")

    errors_loading  = False
    df_students = df_schools = None

    print("\n  [1/4] Loading Student_Info...")
    if not os.path.exists(student_path):
        print(f"  [✗]  File not found: {student_path}")
        errors_loading = True
    else:
        df_students = load_student_info(student_path)
        print(f"  [✓]  {len(df_students)} student records loaded.")

    print("\n  [2/4] Loading School_Info...")
    if not os.path.exists(school_path):
        print(f"  [✗]  File not found: {school_path}")
        errors_loading = True
    else:
        df_schools = load_school_info(school_path)
        print(f"  [✓]  {len(df_schools)} school records loaded.")

    if errors_loading:
        print("\n  ❌  Cannot proceed — fix file paths above.")
        return False

    school_codes   = set(df_schools["School_Code"].astype(str).str.strip())
    total_students = len(df_students)

    print("\n  [3/4] Validating Student_Info...")
    student_report = validate_students(df_students, school_codes)
    print_report("Student_Info Validation", student_report)

    print("\n  [4/4] Validating School_Info...")
    school_report = validate_schools(df_schools, total_students)
    print_report("School_Info Validation", school_report)

    all_passed = student_report.passed and school_report.passed

    print(f"\n{'═'*65}")
    if all_passed:
        total_warnings = len(student_report.warnings) + len(school_report.warnings)
        print(f"  ✅  VALIDATION PASSED — ready to proceed to Step 4 (Allocation)")
        if total_warnings:
            print(f"  ⚠️   {total_warnings} warning(s) — review above before proceeding")
    else:
        total_errors = len(student_report.errors) + len(school_report.errors)
        print(f"  ❌  VALIDATION FAILED — {total_errors} error(s) must be fixed first")
        print(f"      Fix the issues above, save the Excel files, and re-run this script.")
    print(f"{'═'*65}\n")

    return all_passed


# ════════════════════════════════════════════════════════════════════
# Entry point
# ════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    STUDENT_FILE = "data/Student_Info_Template.xlsx"
    SCHOOL_FILE  = "data/School_Info_Template.xlsx"
    passed = run_validation(STUDENT_FILE, SCHOOL_FILE)
    sys.exit(0 if passed else 1)
