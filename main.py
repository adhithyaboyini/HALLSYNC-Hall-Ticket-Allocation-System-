"""
main.py  —  HallSync Master Runner

Runs the entire pipeline in one command:

    python main.py

Steps executed:
    Step 0  →  Aggregate all school files from data/inbox/
    Step 3  →  Validate master student + school data
    Step 4  →  Run allocation engine
    Step 5  →  Generate hall ticket PDFs

Prerequisites:
    • data/inbox/          — drop all school Excel files here
    • data/School_Info_Template.xlsx  — filled by admin
"""

import os, sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from step0_aggregator  import run             as run_aggregator
from step3_validation  import run_validation
from step4_allocation  import run_allocation
from step5_hall_tickets import run            as run_tickets

# ── Paths ─────────────────────────────────────────────────────────────────────
STUDENT_MASTER = os.path.join(BASE_DIR, "data", "Student_Info_Master.xlsx")
SCHOOL_FILE    = os.path.join(BASE_DIR, "data", "School_Info_Template.xlsx")
ALLOC_OUTPUT   = os.path.join(BASE_DIR, "output", "allocation_result.xlsx")
TICKETS_DIR    = os.path.join(BASE_DIR, "output", "hall_tickets")


def banner(msg: str, char: str = "═"):
    line = char * 65
    print(f"\n{line}")
    print(f"  {msg}")
    print(f"{line}")


def main():
    banner("🏫  HallSync — Full Pipeline Runner")

    # ── Step 0: Aggregate ─────────────────────────────────────────────────────
    banner("STEP 0 / 4  —  Data Aggregation")
    run_aggregator()

    if not os.path.exists(STUDENT_MASTER):
        print("\n  ❌  Master database not created. "
              "Check inbox folder and re-run.")
        sys.exit(1)

    # ── Step 3: Validate ──────────────────────────────────────────────────────
    banner("STEP 1 / 4  —  Validation")
    passed = run_validation(STUDENT_MASTER, SCHOOL_FILE)
    if not passed:
        print("\n  ❌  Validation failed. Fix errors above and re-run.")
        sys.exit(1)

    # ── Step 4: Allocate ──────────────────────────────────────────────────────
    banner("STEP 2 / 4  —  Allocation Engine")
    run_allocation(STUDENT_MASTER, SCHOOL_FILE, ALLOC_OUTPUT)

    # ── Step 5: Hall Tickets ──────────────────────────────────────────────────
    banner("STEP 3 / 4  —  Hall Ticket PDF Generation")
    run_tickets(ALLOC_OUTPUT, TICKETS_DIR)

    # ── Done ──────────────────────────────────────────────────────────────────
    banner("✅  HallSync Pipeline Complete!")
    print(f"  Allocation result  →  {ALLOC_OUTPUT}")
    print(f"  Hall ticket PDFs   →  {TICKETS_DIR}/")
    print(f"  Submission report  →  data/Submission_Status_Report.xlsx\n")


if __name__ == "__main__":
    main()
