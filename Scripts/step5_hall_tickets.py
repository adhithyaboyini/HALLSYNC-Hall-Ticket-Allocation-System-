"""
step5_hall_tickets.py  —  HallSync Step 5: Hall Ticket PDF Generator

Layout  : 4 hall tickets per A4 page (2 columns x 2 rows)
Output  : One PDF per centre  →  output/hall_tickets/<Centre_Code>.pdf
ID      : YYDDCCSS  (8 digits, no dashes)
            YY = year last 2 digits
            DD = district sequential number (01, 02 ...)
            CC = centre sequential number (01, 02 ...)
            SS = seat number (01-99)

Input   : output/allocation_result.xlsx  (from step4_allocation.py)
"""

import pandas as pd
import os, sys, hashlib, random as _rnd
from datetime import datetime
from collections import defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas

# ════════════════════════════════════════════════════════════════════
# Paths & config
# ════════════════════════════════════════════════════════════════════
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
ALLOC_FILE  = os.path.join(BASE_DIR, "output", "allocation_result.xlsx")
OUT_DIR     = os.path.join(BASE_DIR, "output", "hall_tickets")
YEAR        = str(datetime.now().year)[-2:]   # "26"

W, H        = A4           # 595.27 x 841.89 pts
MARGIN      = 10 * mm
COLS        = 2
ROWS        = 2
TICKET_W    = (W - 2 * MARGIN - 4 * mm) / COLS     # width of one ticket
TICKET_H    = (H - 2 * MARGIN - 4 * mm) / ROWS     # height of one ticket

# ════════════════════════════════════════════════════════════════════
# Colours
# ════════════════════════════════════════════════════════════════════
C_NAVY   = colors.HexColor("#1A3C6E")
C_BLUE   = colors.HexColor("#2E6DA4")
C_GOLD   = colors.HexColor("#FFC107")
C_LIGHT  = colors.HexColor("#EEF2F9")
C_GREEN  = colors.HexColor("#0D5C2F")
C_WHITE  = colors.white
C_GREY   = colors.HexColor("#555555")
C_BLACK  = colors.black


# ════════════════════════════════════════════════════════════════════
# QR matrix — pure Python, deterministic
# ════════════════════════════════════════════════════════════════════

def _qr_matrix(data: str, size: int = 21) -> list:
    h   = int(hashlib.md5(data.encode()).hexdigest(), 16)
    rng = _rnd.Random(h)
    m   = [[rng.randint(0, 1) for _ in range(size)] for _ in range(size)]

    def finder(r, c):
        pat = [[1,1,1,1,1,1,1],[1,0,0,0,0,0,1],[1,0,1,1,1,0,1],
               [1,0,1,1,1,0,1],[1,0,1,1,1,0,1],[1,0,0,0,0,0,1],
               [1,1,1,1,1,1,1]]
        for dr in range(7):
            for dc in range(7):
                if 0 <= r+dr < size and 0 <= c+dc < size:
                    m[r+dr][c+dc] = pat[dr][dc]
    finder(0, 0); finder(0, size-7); finder(size-7, 0)
    return m


def draw_qr(c, data: str, x: float, y: float, size: float = 18*mm):
    mat  = _qr_matrix(data)
    n    = len(mat)
    cell = size / n
    c.setFillColor(C_WHITE)
    c.rect(x, y, size, size, fill=1, stroke=0)
    c.setFillColor(C_BLACK)
    for r in range(n):
        for col in range(n):
            if mat[r][col]:
                c.rect(x + col*cell, y + (n-1-r)*cell, cell, cell,
                       fill=1, stroke=0)
    c.setStrokeColor(C_NAVY)
    c.setLineWidth(0.3)
    c.rect(x, y, size, size, fill=0, stroke=1)


# ════════════════════════════════════════════════════════════════════
# ID generator
# ════════════════════════════════════════════════════════════════════

def make_ticket_id(year: str, dist_seq: int,
                   centre_seq: int, seat: int) -> str:
    """
    YYDDCCSS — 8 digits, no dashes.
    YY = year, DD = district seq, CC = centre seq, SS = seat
    All zero-padded to their width.
    """
    yy = year.zfill(2)[-2:]
    dd = str(dist_seq).zfill(2)
    cc = str(centre_seq).zfill(2)
    ss = str(seat).zfill(2)
    return f"{yy}{dd}{cc}{ss}"


# ════════════════════════════════════════════════════════════════════
# Single ticket renderer
# ════════════════════════════════════════════════════════════════════

def draw_ticket(c, row: dict, ticket_id: str,
                tx: float, ty: float):
    """
    Draw one hall ticket inside a box of TICKET_W x TICKET_H
    at bottom-left (tx, ty).
    """
    TW = TICKET_W
    TH = TICKET_H
    PAD = 3 * mm

    # ── Outer border ──────────────────────────────────────────────
    c.setStrokeColor(C_NAVY)
    c.setLineWidth(1.2)
    c.roundRect(tx, ty, TW, TH, 3*mm, fill=0, stroke=1)

    # ── Top gold accent strip ─────────────────────────────────────
    c.setFillColor(C_GOLD)
    c.rect(tx, ty + TH - 3*mm, TW, 3*mm, fill=1, stroke=0)

    # ── Header band ───────────────────────────────────────────────
    HDR_H = 14 * mm
    c.setFillColor(C_NAVY)
    c.rect(tx, ty + TH - 3*mm - HDR_H, TW, HDR_H, fill=1, stroke=0)

    # Header text
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(tx + TW/2,
                        ty + TH - 3*mm - HDR_H + 8*mm,
                        "HALL TICKET — ANNUAL EXAMINATION " + f"20{YEAR}")
    c.setFont("Helvetica", 6.5)
    c.setFillColor(C_GOLD)
    c.drawCentredString(tx + TW/2,
                        ty + TH - 3*mm - HDR_H + 3.5*mm,
                        "HallSync Examination Management System")

    # ── Ticket ID band ────────────────────────────────────────────
    ID_Y = ty + TH - 3*mm - HDR_H - 7*mm
    c.setFillColor(C_LIGHT)
    c.rect(tx, ID_Y, TW, 7*mm, fill=1, stroke=0)
    c.setFillColor(C_NAVY)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(tx + PAD, ID_Y + 4.5*mm, "HALL TICKET ID")
    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(C_BLUE)
    c.drawString(tx + PAD + 26*mm, ID_Y + 3.5*mm, ticket_id)

    # QR code (right side of ID band, spans into body)
    QR_SIZE = 22 * mm
    QR_X    = tx + TW - QR_SIZE - PAD
    QR_Y    = ty + TH - 3*mm - HDR_H - 7*mm - QR_SIZE - PAD

    # ── Body area ─────────────────────────────────────────────────
    BODY_TOP = ID_Y - PAD
    BODY_BOT = ty + PAD + 8*mm   # reserve bottom for footer

    draw_qr(c, ticket_id, QR_X, QR_Y, QR_SIZE)

    # ── Student details ───────────────────────────────────────────
    def kv(label, value, x, y, lw=22*mm):
        c.setFont("Helvetica-Bold", 6.5)
        c.setFillColor(C_GREY)
        c.drawString(x, y, label)
        c.setFont("Helvetica", 7.5)
        c.setFillColor(C_BLACK)
        # Truncate long values to fit
        max_chars = int((TW - lw - QR_SIZE - PAD*3) / 3.2)
        val_str   = str(value)[:max_chars] if len(str(value)) > max_chars else str(value)
        c.drawString(x + lw, y, val_str)

    field_x = tx + PAD
    field_y = BODY_TOP

    fields = [
        ("Student Name :", row.get("Student_Name", "")),
        ("School       :", f"{row.get('School_Name','')} ({row.get('School_Code','')})"),
        ("Centre       :", row.get("Centre_Name", "")),
        ("Address      :", row.get("Centre_District", "")),
        ("Hall  /  Seat:", f"{row.get('Hall', '')}  —  {row.get('Seat_Display', '')}"),
    ]

    line_h = 6 * mm
    for label, value in fields:
        kv(label, value, field_x, field_y)
        field_y -= line_h

    # ── Divider line ──────────────────────────────────────────────
    c.setStrokeColor(colors.HexColor("#CCCCCC"))
    c.setLineWidth(0.3)
    c.line(tx + PAD, ty + PAD + 8*mm,
           tx + TW - PAD, ty + PAD + 8*mm)

    # ── Footer ────────────────────────────────────────────────────
    c.setFont("Helvetica-Oblique", 5.5)
    c.setFillColor(C_GREY)
    c.drawString(tx + PAD, ty + PAD + 4.5*mm,
                 "This is a computer-generated hall ticket. "
                 "Produce this at the exam centre.")
    c.drawString(tx + PAD, ty + PAD + 1.5*mm,
                 "Bring school ID card. Report 30 mins before exam.")

    # ── Photo box ─────────────────────────────────────────────────
    PH_W = 14 * mm; PH_H = 17 * mm
    PH_X = tx + TW - PH_W - PAD
    PH_Y = BODY_TOP - PH_H
    c.setFillColor(C_WHITE)
    c.setStrokeColor(C_NAVY)
    c.setLineWidth(0.5)
    c.rect(PH_X, PH_Y, PH_W, PH_H, fill=1, stroke=1)
    c.setFont("Helvetica", 5)
    c.setFillColor(C_GREY)
    c.drawCentredString(PH_X + PH_W/2, PH_Y + PH_H/2 + 1*mm, "PHOTO")
    c.drawCentredString(PH_X + PH_W/2, PH_Y + PH_H/2 - 3*mm, "PASTE HERE")

    # ── Signature line ────────────────────────────────────────────
    sig_y = ty + PAD + 9.5*mm
    sig_x = tx + PAD
    c.setStrokeColor(C_GREY)
    c.setLineWidth(0.4)
    c.line(sig_x, sig_y, sig_x + 30*mm, sig_y)
    c.setFont("Helvetica", 5.5)
    c.setFillColor(C_GREY)
    c.drawString(sig_x + 3*mm, sig_y + 1*mm, "Signature")


# ════════════════════════════════════════════════════════════════════
# Build sequence maps
# ════════════════════════════════════════════════════════════════════

def build_sequence_maps(df: pd.DataFrame) -> tuple[dict, dict]:
    """
    Assign sequential district and centre numbers at runtime.
    Returns:
        dist_map   : district_code  → 2-digit str ("01","02"...)
        centre_map : centre_code    → 2-digit str
    """
    dist_map   = {}
    centre_map = {}
    d_ctr = 1
    c_ctr = 1

    for _, row in df.sort_values(["Centre_District",
                                  "Centre_Code"]).iterrows():
        dist   = str(row.get("Centre_District", "")).strip().upper()
        centre = str(row.get("Centre_Code",     "")).strip().upper()

        if dist not in dist_map:
            dist_map[dist] = str(d_ctr).zfill(2)
            d_ctr += 1
        if centre not in centre_map:
            centre_map[centre] = str(c_ctr).zfill(2)
            c_ctr += 1

    return dist_map, centre_map


# ════════════════════════════════════════════════════════════════════
# Generate one PDF per centre
# ════════════════════════════════════════════════════════════════════

def generate_centre_pdf(centre_code: str,
                        centre_name: str,
                        students: list[dict],
                        dist_map: dict,
                        centre_map: dict,
                        out_dir: str) -> str:
    """
    Generate a PDF for all students at one centre.
    Layout: 4 tickets per page (2 cols x 2 rows).
    """
    safe_name = centre_code.replace("/", "_").replace(" ", "_")
    filepath  = os.path.join(out_dir, f"{safe_name}_HallTickets.pdf")

    c = rl_canvas.Canvas(filepath, pagesize=A4)
    c.setTitle(f"Hall Tickets — {centre_name}")

    # Ticket positions on page (bottom-left of each slot)
    def slot_pos(idx: int) -> tuple[float, float]:
        col = idx % COLS
        row = idx // COLS         # 0 = top row, 1 = bottom row
        tx  = MARGIN + col * (TICKET_W + 4*mm)
        # row 0 → top half, row 1 → bottom half
        ty  = H - MARGIN - (row + 1) * TICKET_H - row * 4*mm
        return tx, ty

    page_idx    = 0
    tickets_pp  = COLS * ROWS   # 4

    for i, student in enumerate(students):
        slot = i % tickets_pp

        if slot == 0 and i > 0:
            # Draw page number before turning page
            _draw_page_number(c, page_idx + 1, centre_name)
            c.showPage()
            page_idx += 1

        tx, ty = slot_pos(slot)

        # Build ticket ID
        dist_seq   = int(dist_map.get(
            str(student.get("Centre_District","")).strip().upper(), "01"))
        centre_seq = int(centre_map.get(
            str(student.get("Centre_Code","")).strip().upper(), "01"))
        seat_num   = int(str(student.get("Seat_Number", 1)))
        ticket_id  = make_ticket_id(YEAR, dist_seq, centre_seq, seat_num)

        draw_ticket(c, student, ticket_id, tx, ty)

    # Last page number
    _draw_page_number(c, page_idx + 1, centre_name)

    c.save()
    return filepath


def _draw_page_number(c, page_num: int, centre_name: str):
    c.setFont("Helvetica", 7)
    c.setFillColor(C_GREY)
    c.drawCentredString(W/2, 6*mm,
                        f"{centre_name}  |  Page {page_num}")
    c.setStrokeColor(colors.HexColor("#DDDDDD"))
    c.setLineWidth(0.3)
    c.line(MARGIN, 9*mm, W - MARGIN, 9*mm)


# ════════════════════════════════════════════════════════════════════
# Main runner
# ════════════════════════════════════════════════════════════════════

def run(alloc_path: str = ALLOC_FILE, out_dir: str = OUT_DIR):

    print(f"\n{'═'*65}")
    print(f"  HallSync — Step 5: Hall Ticket PDF Generator")
    print(f"{'═'*65}")

    # ── Load allocation result ────────────────────────────────────
    if not os.path.exists(alloc_path):
        print(f"\n  ❌  Allocation file not found: {alloc_path}")
        print(f"      Run step4_allocation.py first.")
        sys.exit(1)

    df = pd.read_excel(alloc_path, dtype=str)
    df.columns = df.columns.str.strip()
    df["Seat_Number"] = pd.to_numeric(df["Seat_Number"],
                                      errors="coerce").fillna(1).astype(int)
    print(f"\n  [✓] Loaded {len(df)} allocated students")

    os.makedirs(out_dir, exist_ok=True)

    # ── Build sequence maps ───────────────────────────────────────
    dist_map, centre_map = build_sequence_maps(df)
    print(f"\n  District map  : {dist_map}")
    print(f"  Centre map    : {centre_map}")

    # ── Group by centre ───────────────────────────────────────────
    centre_groups = defaultdict(list)
    for _, row in df.sort_values(["Centre_Code",
                                   "Hall", "Seat_Number"]).iterrows():
        ccode = str(row.get("Centre_Code", "")).strip().upper()
        centre_groups[ccode].append(row.to_dict())

    # ── Generate one PDF per centre ───────────────────────────────
    print(f"\n  Generating PDFs for {len(centre_groups)} centre(s)...\n")
    generated = []

    for ccode, students in centre_groups.items():
        cname = students[0].get("Centre_Name", ccode)
        pages = -(-len(students) // (COLS * ROWS))   # ceiling division
        fp    = generate_centre_pdf(ccode, cname, students,
                                    dist_map, centre_map, out_dir)
        generated.append(fp)
        print(f"  ✅  {cname:<40} "
              f"{len(students):>3} students  "
              f"{pages:>2} page(s)  →  {os.path.basename(fp)}")

    # ── Summary ───────────────────────────────────────────────────
    print(f"\n{'═'*65}")
    print(f"  ✅  {len(generated)} PDF(s) generated → {out_dir}/")
    print(f"  Total hall tickets : {len(df)}")
    print(f"{'═'*65}\n")

    return generated


if __name__ == "__main__":
    run()
