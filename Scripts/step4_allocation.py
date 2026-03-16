"""
step4_allocation.py  —  HallSync Step 4: Allocation Engine

Rules enforced (in priority order):
  R1  Student cannot be allocated to their own school's centre
  R2  Prefer centres in the same district first, then any available centre
  R3  No two consecutive seats from the same school (interleave by school)
  R4  Best-effort gender alternation M → F → M → F seat by seat
       - If opposite gender unavailable, place same gender
  R5  Halls auto-calculated at SEATS_PER_HALL (default 30) seats per hall
  R6  Fill one hall completely before moving to next
"""

import pandas as pd
import os, sys
from collections import defaultdict, deque
from dataclasses import dataclass, field
from typing import Optional


# ════════════════════════════════════════════════════════════════════
# Configuration
# ════════════════════════════════════════════════════════════════════
SEATS_PER_HALL = 30   # auto-hall size


# ════════════════════════════════════════════════════════════════════
# Data structures
# ════════════════════════════════════════════════════════════════════

@dataclass
class Centre:
    school_code:  str
    school_name:  str
    district:     str
    capacity:     int
    halls:        list = field(default_factory=list)   # list of Hall
    allocated:    int  = 0

    def is_full(self):
        return self.allocated >= self.capacity

    def remaining(self):
        return self.capacity - self.allocated


@dataclass
class Hall:
    centre_code:  str
    hall_label:   str          # H1, H2, ...
    capacity:     int
    seats:        list = field(default_factory=list)   # filled AllocationRecord
    allocated:    int  = 0

    def is_full(self):
        return self.allocated >= self.capacity

    def remaining(self):
        return self.capacity - self.allocated


@dataclass
class AllocationRecord:
    student_id:       str
    student_name:     str
    father_name:      str
    mother_name:      str
    gender:           str
    dob:              str
    language_1:       str
    language_2:       str
    medium:           str
    subjects:         str
    total_subjects:   str
    school_name:      str
    school_code:      str
    area:             str
    district_code:    str
    centre_code:      str
    centre_name:      str
    centre_district:  str
    hall_label:       str
    seat_number:      int
    seat_display:     str      # e.g. "H1 - 01"


# ════════════════════════════════════════════════════════════════════
# Loaders  (reuse same logic as validator)
# ════════════════════════════════════════════════════════════════════

def load_students(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Student_Info", header=4, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    df = df[~df["Student_ID"].astype(str).str.strip().str.lower()
              .isin(["s001", "nan", "none", ""])]
    return df.reset_index(drop=True)


def load_schools(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="School_Info", header=4, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    df = df[~df["School_Code"].astype(str).str.strip().str.lower()
              .isin(["nan", "none", "", "totals"])]
    df = df[~df["School_Code"].astype(str).str.startswith("=")]
    return df.reset_index(drop=True)


# ════════════════════════════════════════════════════════════════════
# Build centre pool
# ════════════════════════════════════════════════════════════════════

def build_centres(df_schools: pd.DataFrame) -> dict[str, Centre]:
    """Build a dict of Centre objects keyed by school_code (uppercase)."""
    centres = {}
    for _, row in df_schools.iterrows():
        if str(row.get("Can_Be_Centre", "")).strip().capitalize() != "Yes":
            continue
        try:
            cap = int(float(str(row.get("Capacity", "0")).strip()))
        except ValueError:
            cap = 0
        if cap <= 0:
            continue

        code     = str(row["School_Code"]).strip().upper()
        district = str(row.get("District_Code", "")).strip().upper()

        # Build halls
        halls       = []
        hall_num    = 1
        remaining   = cap
        while remaining > 0:
            hall_cap = min(SEATS_PER_HALL, remaining)
            halls.append(Hall(
                centre_code=code,
                hall_label=f"H{hall_num}",
                capacity=hall_cap,
            ))
            remaining -= hall_cap
            hall_num  += 1

        centres[code] = Centre(
            school_code=code,
            school_name=str(row["School_Name"]).strip(),
            district=district,
            capacity=cap,
            halls=halls,
        )
    return centres


# ════════════════════════════════════════════════════════════════════
# Smart seat queue — R3 + R4 combined
# ════════════════════════════════════════════════════════════════════

class SeatQueue:
    """
    Manages ordered seating for one hall.

    R3: No two consecutive seats from the same school.
        Achieved by interleaving students from different schools.
    R4: Best-effort gender alternation M → F → M → F.
        Achieved by picking from male/female buckets alternately.

    Strategy:
      - Maintain two buckets per school: males and females
      - At each seat, pick the gender we want (alternating)
      - Among candidates of that gender, pick from a school
        different from the last placed school
      - If no opposite gender available, pick same gender
        (different school still enforced if possible)
    """

    def __init__(self):
        # school_code → {"M": deque([student, ...]), "F": deque([...])}
        self._buckets: dict[str, dict[str, deque]] = defaultdict(
            lambda: {"M": deque(), "F": deque(), "O": deque()}
        )
        self._total    = 0
        self._last_school = None
        self._want_gender = "M"   # start with Male

    def _gender_key(self, gender: str) -> str:
        g = gender.strip().upper()
        if g in ("MALE", "M"):      return "M"
        if g in ("FEMALE", "F"):    return "F"
        return "O"

    def add(self, student: dict):
        scode = str(student.get("School_Code", "")).strip().upper()
        gkey  = self._gender_key(str(student.get("Gender", "")))
        self._buckets[scode][gkey].append(student)
        self._total += 1

    def _pick_from_gender(self, gender_key: str,
                          exclude_school: Optional[str]) -> Optional[dict]:
        """
        Pick a student of given gender, preferring a school
        different from exclude_school.
        Sorts by bucket size descending so larger schools get spread first.
        """
        sorted_schools = sorted(
            self._buckets.items(),
            key=lambda x: -(len(x[1]["M"]) + len(x[1]["F"]) + len(x[1]["O"]))
        )
        # First pass: different school, largest bucket first
        for scode, buckets in sorted_schools:
            if scode == exclude_school:
                continue
            if buckets[gender_key]:
                student = buckets[gender_key].popleft()
                self._total -= 1
                return student
        # Second pass: same school unavoidable
        for scode, buckets in sorted_schools:
            if buckets[gender_key]:
                student = buckets[gender_key].popleft()
                self._total -= 1
                return student
        return None

    def pop_next(self) -> Optional[dict]:
        """
        Pop the next student following R3 + R4.
        Returns None if queue is empty.
        """
        if self._total == 0:
            return None

        # Try wanted gender first (R4)
        student = self._pick_from_gender(self._want_gender, self._last_school)

        # Fallback: try other genders if wanted not available (R4 best-effort)
        if student is None:
            for gkey in ("M", "F", "O"):
                if gkey == self._want_gender:
                    continue
                student = self._pick_from_gender(gkey, self._last_school)
                if student:
                    break

        if student is None:
            return None

        gkey = self._gender_key(str(student.get("Gender", "")))
        # Flip wanted gender for next seat (R4)
        self._want_gender = "F" if self._want_gender == "M" else "M"
        # Track last school (R3)
        self._last_school = str(student.get("School_Code", "")).strip().upper()
        return student

    def __len__(self):
        return self._total


# ════════════════════════════════════════════════════════════════════
# Allocation engine
# ════════════════════════════════════════════════════════════════════

def allocate(df_students: pd.DataFrame,
             centres: dict[str, Centre]) -> tuple[list[AllocationRecord], list[str]]:
    """
    Main allocation function.

    Returns:
        records      : list of AllocationRecord for every allocated student
        unallocated  : list of Student_IDs that could not be placed
    """
    records     = []
    unallocated = []

    # ── Pre-process students ──────────────────────────────────────────────────
    # Sort: by district first (so same-district students processed together),
    # then by school, then by student_id (deterministic)
    df = df_students.copy()
    df["_district_upper"] = df["District_Code"].astype(str).str.strip().str.upper()
    df["_school_upper"]   = df["School_Code"].astype(str).str.strip().str.upper()
    df = df.sort_values(["_district_upper", "_school_upper", "Student_ID"]) \
           .reset_index(drop=True)

    # Group students into per-centre SeatQueues
    # Each centre gets a SeatQueue per hall
    # We fill halls sequentially, so we build one global SeatQueue
    # and distribute as we fill each hall

    # ── Assign each student to a centre (R1, R2) ─────────────────────────────
    # Build a mapping: student → assigned_centre
    centre_assignments: dict[str, str] = {}   # student_id → centre_code

    # For each student, find the best centre:
    #   Priority 1: same district, not own school
    #   Priority 2: any centre, not own school
    centre_list = list(centres.values())

    for _, student in df.iterrows():
        sid         = str(student["Student_ID"]).strip()
        own_school  = str(student.get("School_Code", "")).strip().upper()
        s_district  = str(student.get("District_Code", "")).strip().upper()

        # Filter eligible centres (R1: not own school)
        eligible = [c for c in centre_list if c.school_code != own_school and not c.is_full()]

        if not eligible:
            unallocated.append(sid)
            continue

        # R2: prefer same district
        same_dist = [c for c in eligible if c.district == s_district]
        chosen    = (same_dist or eligible)

        # Among eligible, pick centre with most remaining capacity (balance load)
        chosen.sort(key=lambda c: -c.remaining())
        centre = chosen[0]

        centre_assignments[sid] = centre.school_code
        centre.allocated += 1   # reserve the seat now to avoid overflow

    # Reset allocated count — we'll recount during hall filling
    for c in centres.values():
        c.allocated = 0

    # ── Build per-centre SeatQueues ───────────────────────────────────────────
    centre_queues: dict[str, SeatQueue] = {code: SeatQueue() for code in centres}

    for _, student in df.iterrows():
        sid = str(student["Student_ID"]).strip()
        if sid not in centre_assignments:
            continue
        ccode = centre_assignments[sid]
        centre_queues[ccode].add(student.to_dict())

    # ── Fill halls seat by seat ───────────────────────────────────────────────
    for ccode, centre in centres.items():
        queue = centre_queues.get(ccode)
        if not queue or len(queue) == 0:
            continue

        for hall in centre.halls:
            seat_num = 1
            while not hall.is_full() and len(queue) > 0:
                student = queue.pop_next()
                if student is None:
                    break

                rec = AllocationRecord(
                    student_id      = str(student.get("Student_ID", "")).strip(),
                    student_name    = str(student.get("Student_Name", "")).strip(),
                    father_name     = str(student.get("Father_Name", "")).strip(),
                    mother_name     = str(student.get("Mother_Name", "")).strip(),
                    gender          = str(student.get("Gender", "")).strip(),
                    dob             = str(student.get("Date_of_Birth", "")).strip(),
                    language_1      = str(student.get("Language_1", "")).strip(),
                    language_2      = str(student.get("Language_2", "")).strip(),
                    medium          = str(student.get("Medium_of_Instruction", "")).strip(),
                    subjects        = str(student.get("Subjects", "")).strip(),
                    total_subjects  = str(student.get("Total_Subjects", "")).strip(),
                    school_name     = str(student.get("School_Name", "")).strip(),
                    school_code     = str(student.get("School_Code", "")).strip(),
                    area            = str(student.get("Area", "")).strip(),
                    district_code   = str(student.get("District_Code", "")).strip(),
                    centre_code     = centre.school_code,
                    centre_name     = centre.school_name,
                    centre_district = centre.district,
                    hall_label      = hall.hall_label,
                    seat_number     = seat_num,
                    seat_display    = f"{hall.hall_label} - {seat_num:02d}",
                )
                hall.seats.append(rec)
                hall.allocated += 1
                centre.allocated += 1
                records.append(rec)
                seat_num += 1

        # Any remaining in queue after all halls filled → unallocated
        while len(queue) > 0:
            student = queue.pop_next()
            if student:
                unallocated.append(str(student.get("Student_ID", "")).strip())

    return records, unallocated


# ════════════════════════════════════════════════════════════════════
# Rule verification
# ════════════════════════════════════════════════════════════════════

def verify_rules(records: list[AllocationRecord],
                 centres: dict[str, Centre]) -> dict:
    """
    Post-allocation rule verification.
    Returns a dict of rule → pass/fail with details.
    """
    results = {}

    # R1: no student at own school
    r1_violations = [r for r in records if r.school_code.upper() == r.centre_code.upper()]
    results["R1_own_school"] = {
        "passed": len(r1_violations) == 0,
        "violations": len(r1_violations),
        "details": [(r.student_id, r.school_code) for r in r1_violations[:5]]
    }

    # R3: consecutive same-school check per hall
    r3_violations = 0
    for centre in centres.values():
        for hall in centre.halls:
            for i in range(1, len(hall.seats)):
                if (hall.seats[i].school_code.upper() ==
                        hall.seats[i-1].school_code.upper()):
                    r3_violations += 1
    results["R3_same_school_consecutive"] = {
        "passed": r3_violations == 0,
        "violations": r3_violations,
    }

    # R4: gender alternation score
    gender_pairs = 0
    gender_alternates = 0
    for centre in centres.values():
        for hall in centre.halls:
            for i in range(1, len(hall.seats)):
                gender_pairs += 1
                if hall.seats[i].gender.upper() != hall.seats[i-1].gender.upper():
                    gender_alternates += 1
    alt_pct = round(gender_alternates / gender_pairs * 100, 1) if gender_pairs > 0 else 0
    results["R4_gender_alternation"] = {
        "passed": alt_pct >= 60,   # 60%+ is considered good
        "alternation_pct": alt_pct,
        "alternating_pairs": gender_alternates,
        "total_pairs": gender_pairs,
    }

    return results


# ════════════════════════════════════════════════════════════════════
# Summary printer
# ════════════════════════════════════════════════════════════════════

def print_summary(records: list[AllocationRecord],
                  unallocated: list[str],
                  centres: dict[str, Centre],
                  rule_results: dict):

    print(f"\n{'═'*65}")
    print(f"  HallSync — Allocation Summary")
    print(f"{'═'*65}")
    print(f"  Total Allocated  : {len(records)}")
    print(f"  Unallocated      : {len(unallocated)}")
    if unallocated:
        print(f"  Unallocated IDs  : {', '.join(unallocated[:10])}")

    print(f"\n  Centre Breakdown:")
    print(f"  {'Centre':<35} {'District':<15} {'Halls':>6} {'Seats':>6}")
    print(f"  {'─'*62}")
    for code, centre in centres.items():
        if centre.allocated == 0:
            continue
        halls_used = sum(1 for h in centre.halls if h.allocated > 0)
        print(f"  {centre.school_name:<35} {centre.district:<15} "
              f"{halls_used:>6} {centre.allocated:>6}")

    print(f"\n  Rule Verification:")
    print(f"  {'─'*50}")
    r = rule_results
    icon = lambda p: "✅" if p else "❌"

    print(f"  {icon(r['R1_own_school']['passed'])}  R1 Own-school violations : "
          f"{r['R1_own_school']['violations']}")
    print(f"  {icon(r['R3_same_school_consecutive']['passed'])}  R3 Consecutive same-school: "
          f"{r['R3_same_school_consecutive']['violations']}")
    print(f"  {icon(r['R4_gender_alternation']['passed'])}  R4 Gender alternation    : "
          f"{r['R4_gender_alternation']['alternation_pct']}%  "
          f"({r['R4_gender_alternation']['alternating_pairs']} / "
          f"{r['R4_gender_alternation']['total_pairs']} pairs)")
    print(f"{'═'*65}\n")


# ════════════════════════════════════════════════════════════════════
# Save allocation output
# ════════════════════════════════════════════════════════════════════

def save_allocation(records: list[AllocationRecord],
                    output_path: str):
    rows = []
    for r in records:
        rows.append({
            "Student_ID":       r.student_id,
            "Student_Name":     r.student_name,
            "Father_Name":      r.father_name,
            "Mother_Name":      r.mother_name,
            "Gender":           r.gender,
            "Date_of_Birth":    r.dob,
            "Language_1":       r.language_1,
            "Language_2":       r.language_2,
            "Medium":           r.medium,
            "Subjects":         r.subjects,
            "School_Name":      r.school_name,
            "School_Code":      r.school_code,
            "Area":             r.area,
            "District_Code":    r.district_code,
            "Centre_Code":      r.centre_code,
            "Centre_Name":      r.centre_name,
            "Centre_District":  r.centre_district,
            "Hall":             r.hall_label,
            "Seat_Number":      r.seat_number,
            "Seat_Display":     r.seat_display,
        })
    df = pd.DataFrame(rows)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False)
    print(f"  [✓] Allocation saved → {output_path}")
    return df


# ════════════════════════════════════════════════════════════════════
# Main
# ════════════════════════════════════════════════════════════════════

def run_allocation(student_path: str,
                   school_path: str,
                   output_path: str) -> list[AllocationRecord]:

    print(f"\n{'═'*65}")
    print(f"  HallSync — Step 4: Allocation Engine")
    print(f"{'═'*65}")

    print("\n  [1/5] Loading data...")
    df_students = load_students(student_path)
    df_schools  = load_schools(school_path)
    print(f"        Students : {len(df_students)}")
    print(f"        Schools  : {len(df_schools)}")

    print("\n  [2/5] Building centre pool...")
    centres = build_centres(df_schools)
    total_cap = sum(c.capacity for c in centres.values())
    print(f"        Centres  : {len(centres)}")
    print(f"        Total capacity : {total_cap}")
    for code, c in centres.items():
        print(f"          {c.school_name:<35} capacity={c.capacity}  "
              f"halls={len(c.halls)}")

    print("\n  [3/5] Running allocation algorithm...")
    records, unallocated = allocate(df_students, centres)
    print(f"        Allocated    : {len(records)}")
    print(f"        Unallocated  : {len(unallocated)}")

    print("\n  [4/5] Verifying rules...")
    rule_results = verify_rules(records, centres)

    print("\n  [5/5] Saving results...")
    save_allocation(records, output_path)

    print_summary(records, unallocated, centres, rule_results)
    return records


# ════════════════════════════════════════════════════════════════════
# Entry point
# ════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    STUDENT_FILE = "data/Student_Info_Template.xlsx"
    SCHOOL_FILE  = "data/School_Info_Template.xlsx"
    OUTPUT_FILE  = "output/allocation_result.xlsx"
    run_allocation(STUDENT_FILE, SCHOOL_FILE, OUTPUT_FILE)
