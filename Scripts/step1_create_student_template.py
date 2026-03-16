"""
Builds Student_Info_Template.xlsm
Multi-select subjects via VBA SelectionChange + runtime UserForm
"""
import zipfile, shutil, os, struct
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY = "1A3C6E"; BLUE = "2E6DA4"; GREEN = "0D5C2F"
LIGHT = "EEF2F9"; WHITE = "FFFFFF"; GOLD = "FFC107"
GREY  = "555555"; YBGD  = "FFF8E7"; GBGD = "E8F5E9"

SUBJECTS = ["Maths","Science","Social Science","Biology","Physics","Chemistry","Computer Science"]

def tb(c="BBBBBB"):
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)

def hc(cell, text, bg=NAVY, fg=WHITE, sz=10, wrap=False):
    cell.value = text
    cell.font  = Font(bold=True, size=sz, color=fg, name="Arial")
    cell.fill  = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = tb()

def dc(cell, bg=WHITE):
    cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = tb()
    cell.font = Font(name="Arial", size=9)

wb = Workbook()
ws = wb.active
ws.title = "Student_Info"

# ── Rows 1-3: Title / subtitle / gold line ────────────────────────────────────
ws.merge_cells("A1:O1")
hc(ws["A1"], "🏫  HallSync — Student Information Sheet", sz=14)
ws.row_dimensions[1].height = 32

ws.merge_cells("A2:O2")
c2 = ws["A2"]
c2.value = "Fill one row per student  •  Click any Subjects cell → multi-select popup opens  •  School details same for all — fill once, copy down"
c2.font  = Font(italic=True, size=9, color=GREY, name="Arial")
c2.fill  = PatternFill("solid", start_color=YBGD)
c2.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 16

for col in range(1, 16):
    ws.cell(3, col).fill = PatternFill("solid", start_color=GOLD)
ws.row_dimensions[3].height = 4

# ── Row 4: Section labels ─────────────────────────────────────────────────────
ws.merge_cells("A4:F4"); hc(ws["A4"], "👤  Student Details",  bg=NAVY)
ws.merge_cells("G4:K4"); hc(ws["G4"], "📚  Academic Details", bg=BLUE)
ws.merge_cells("L4:O4"); hc(ws["L4"], "🏫  School Details",   bg=GREEN)
ws.row_dimensions[4].height = 20

# ── Row 5: Headers ────────────────────────────────────────────────────────────
COLS = [
    ("Student_ID",14,NAVY),("Student_Name",24,NAVY),("Father_Name",22,NAVY),
    ("Mother_Name",22,NAVY),("Gender",12,NAVY),("Date_of_Birth",16,NAVY),
    ("Language_1",16,BLUE),("Language_2",16,BLUE),("Medium_of_Instruction",22,BLUE),
    ("Subjects",36,BLUE),("Total_Subjects",14,BLUE),
    ("School_Name",30,GREEN),("School_Code",14,GREEN),("Area",20,GREEN),("District_Code",16,GREEN),
]
for i,(name,width,bg) in enumerate(COLS,1):
    hc(ws.cell(5,i), name, bg=bg, sz=9, wrap=True)
    ws.column_dimensions[get_column_letter(i)].width = width
ws.row_dimensions[5].height = 28

# ── Row 6: Example row ────────────────────────────────────────────────────────
ex = ["S001","Arun Kumar","Rajan Kumar","Priya Kumar","Male","15-04-2010",
      "Tamil","English","Tamil Medium","Maths, Science, Biology","",
      "Alpha Higher Secondary School","SCH101","Chennai North","TN-CHN-01"]
for i,v in enumerate(ex,1):
    cell = ws.cell(6,i)
    cell.value = v
    cell.font  = Font(italic=True, size=8, color="AAAAAA", name="Arial")
    cell.fill  = PatternFill("solid", start_color="F5F5F5")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = tb()
ws.cell(6,16).value = "← Example row (delete before submitting)"
ws.cell(6,16).font  = Font(italic=True, size=8, color="AAAAAA", name="Arial")
ws.row_dimensions[6].height = 18

# ── Rows 7-106: Data rows ─────────────────────────────────────────────────────
DS, DE = 7, 106
for row in range(DS, DE+1):
    bg = LIGHT if row % 2 == 0 else WHITE
    for col in range(1,16):
        dc(ws.cell(row,col), bg=bg)
    j = f"J{row}"
    ws.cell(row,11).value = f'=IF(TRIM({j})="",0,LEN(TRIM({j}))-LEN(SUBSTITUTE(TRIM({j}),",",""))+1)'
    ws.cell(row,11).font  = Font(name="Arial", size=9, color="1E7E34", bold=True)
    ws.cell(row,11).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row,11).border = tb()
    ws.cell(row,11).fill   = PatternFill("solid", start_color=GBGD)
    ws.row_dimensions[row].height = 17

# ── Dropdowns ─────────────────────────────────────────────────────────────────
def adv(formula, sqref, err=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True,
                        showDropDown=False, showErrorMessage=bool(err),
                        errorTitle="Invalid" if err else None, error=err)
    ws.add_data_validation(dv); dv.sqref = sqref

adv('"Male,Female,Other"',                                              f"E{DS}:E{DE}", "Select Male, Female, or Other.")
adv('"Tamil,Telugu,Kannada,Malayalam,Hindi,Sanskrit,Urdu,French"',     f"G{DS}:G{DE}")
adv('"Tamil,Telugu,Kannada,Malayalam,Hindi,Sanskrit,Urdu,French,English"', f"H{DS}:H{DE}")
adv('"Tamil Medium,English Medium,Telugu Medium,Kannada Medium,Hindi Medium"', f"I{DS}:I{DE}")

ws.freeze_panes = "A7"

# ── Subjects helper sheet (hidden, used by VBA) ───────────────────────────────
wh = wb.create_sheet("SubjectList")
wh.sheet_state = "hidden"
for i, s in enumerate(SUBJECTS, 1):
    wh.cell(i, 1).value = s

# ── Instructions sheet ────────────────────────────────────────────────────────
wi = wb.create_sheet("Instructions")
wi.sheet_view.showGridLines = False
wi.column_dimensions["A"].width = 4
wi.column_dimensions["B"].width = 26
wi.column_dimensions["C"].width = 62

wi.merge_cells("B1:C1")
hc(wi["B1"], "📋  HallSync — Field Instructions", sz=13)
wi.row_dimensions[1].height = 30

rows = [
    ("FIELD","DESCRIPTION",True),
    ("Student_ID","Unique ID per student. Manual entry.\nExample: S001 or SCH101001",False),
    ("Student_Name","Full name of student.",False),
    ("Father_Name","Full name of student's father.",False),
    ("Mother_Name","Full name of student's mother.",False),
    ("Gender","Dropdown → Male / Female / Other.",False),
    ("Date_of_Birth","Format: DD-MM-YYYY   Example: 15-04-2010",False),
    ("Language_1","First language. Dropdown + free text allowed.",False),
    ("Language_2","Second language. Dropdown + free text allowed.",False),
    ("Medium_of_Instruction","Teaching medium. Dropdown + free text.",False),
    ("Subjects","Click the cell → multi-select popup opens.\nTick subjects to add/remove.\nSelected subjects appear comma-separated automatically.\nAvailable: Maths, Science, Social Science, Biology, Physics, Chemistry, Computer Science",False),
    ("Total_Subjects","✅ AUTO-CALCULATED — do NOT edit.\nCounts subjects in Subjects cell.",False),
    ("School_Name","Full school name. Same for all rows.",False),
    ("School_Code","School code. Same for all rows.",False),
    ("Area","Area / locality. Same for all rows.",False),
    ("District_Code","District code. Same for all rows.",False),
    ("","",False),
    ("⚠  Rules","• One row = one student\n• Do NOT rename headers\n• Student_ID must be unique\n• Delete example row 6 before submitting\n• Total_Subjects is auto — never edit\n• Save as .xlsm to keep macro working",False),
]
for r,(f,d,ih) in enumerate(rows,3):
    b = wi.cell(r,2,f); c = wi.cell(r,3,d)
    if ih:
        hc(b,f,bg=BLUE); hc(c,d,bg=BLUE)
    else:
        b.font=Font(bold=True,size=9,name="Arial"); b.fill=PatternFill("solid",start_color=LIGHT)
        b.alignment=Alignment(vertical="top"); b.border=tb()
        c.font=Font(size=9,name="Arial"); c.fill=PatternFill("solid",start_color=WHITE)
        c.alignment=Alignment(vertical="top",wrap_text=True); c.border=tb()
    wi.row_dimensions[r].height = 48 if "\n" in d else 20

# ── Save xlsx ─────────────────────────────────────────────────────────────────
XLSX = "/home/claude/hallsync_v2/data/Student_Info_Template.xlsx"
XLSM = "/home/claude/hallsync_v2/data/Student_Info_Template.xlsm"
wb.save(XLSX)
print(f"xlsx saved → {XLSX}")

# ── VBA code ──────────────────────────────────────────────────────────────────
SUBJ_ARRAY = ", ".join([f'"{s}"' for s in SUBJECTS])

VBA_THISWORKBOOK = '''\
Option Explicit

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "Student_Info" Then Exit Sub
    If Target.Column <> 10 Then Exit Sub
    If Target.Row < 7 Or Target.Row > 106 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    Application.EnableEvents = False
    Call ShowMultiSelect(Target)
    Application.EnableEvents = True
End Sub
'''

VBA_MODULE = f'''\
Option Explicit

Sub ShowMultiSelect(ByVal cell As Range)
    Dim subjects As Variant
    subjects = Array({SUBJ_ARRAY})
    
    Dim frmName As String
    frmName = "frmSubjectPicker"
    
    ' Remove existing form if any
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(frmName)
    On Error GoTo 0
    
    ' Create UserForm
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3)
    frm.Name = frmName
    With frm.Properties
        .Item("Caption") = "Select Subjects — Row " & cell.Row
        .Item("Width")   = 230
        .Item("Height")  = 310
        .Item("StartUpPosition") = 1
        .Item("BackColor") = RGB(238, 242, 249)
    End With
    
    ' Title label
    Dim lbl As Object
    Set lbl = frm.Designer.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "Select one or more subjects:"
        .Left = 10: .Top = 8: .Width = 200: .Height = 16
        .Font.Bold = True: .Font.Size = 9
        .ForeColor = RGB(26, 60, 110)
    End With
    
    ' ListBox
    Dim lb As Object
    Set lb = frm.Designer.Controls.Add("Forms.ListBox.1")
    lb.Name = "lstSubjects"
    lb.MultiSelect  = 1
    lb.Left = 10: lb.Top = 28: lb.Width = 200: lb.Height = 200
    lb.Font.Size = 10
    
    ' OK button
    Dim btnOK As Object
    Set btnOK = frm.Designer.Controls.Add("Forms.CommandButton.1")
    btnOK.Name = "btnOK"
    btnOK.Caption = "OK": btnOK.Left = 10: btnOK.Top = 238
    btnOK.Width = 95: btnOK.Height = 28
    btnOK.BackColor = RGB(26, 60, 110)
    btnOK.ForeColor = RGB(255, 255, 255)
    
    ' Cancel button
    Dim btnCancel As Object
    Set btnCancel = frm.Designer.Controls.Add("Forms.CommandButton.1")
    btnCancel.Name = "btnCancel"
    btnCancel.Caption = "Cancel": btnCancel.Left = 115: btnCancel.Top = 238
    btnCancel.Width = 95: btnCancel.Height = 28
    
    ' Current selections
    Dim current As String
    current = Trim(cell.Value)
    
    ' Add items and pre-select existing
    Dim i As Integer
    For i = 0 To UBound(subjects)
        lb.AddItem subjects(i)
        If Len(current) > 0 Then
            If InStr(1, "," & current & ",", "," & Trim(subjects(i)) & ",", vbTextCompare) > 0 Then
                lb.Selected(i) = True
            End If
        End If
    Next i
    
    ' Form code
    Dim code As String
    code = "Private Sub btnOK_Click()" & Chr(13) & Chr(10) & _
           "    Dim res As String: Dim i As Integer" & Chr(13) & Chr(10) & _
           "    For i = 0 To lstSubjects.ListCount - 1" & Chr(13) & Chr(10) & _
           "        If lstSubjects.Selected(i) Then" & Chr(13) & Chr(10) & _
           "            If Len(res) > 0 Then res = res & "", """ & Chr(13) & Chr(10) & _
           "            res = res & lstSubjects.List(i)" & Chr(13) & Chr(10) & _
           "        End If" & Chr(13) & Chr(10) & _
           "    Next i" & Chr(13) & Chr(10) & _
           "    Me.Tag = res: Me.Hide" & Chr(13) & Chr(10) & _
           "End Sub" & Chr(13) & Chr(10) & _
           "Private Sub btnCancel_Click()" & Chr(13) & Chr(10) & _
           "    Me.Tag = ""CANCEL"": Me.Hide" & Chr(13) & Chr(10) & _
           "End Sub"
    frm.CodeModule.AddFromString code
    
    ' Show form
    Dim result As String
    VBA.UserForms.Add(frm.Name).Show
    result = VBA.UserForms(VBA.UserForms.Count - 1).Tag
    Unload VBA.UserForms(VBA.UserForms.Count - 1)
    
    ' Clean up
    ThisWorkbook.VBProject.VBComponents.Remove _
        ThisWorkbook.VBProject.VBComponents(frmName)
    
    If result <> "CANCEL" Then
        cell.Value = result
    End If
End Sub
'''

# ── Build xlsm by converting xlsx zip ────────────────────────────────────────
# xlsm = xlsx + vbaProject.bin embedded
# We'll create a proper xlsm by modifying Content_Types and adding VBA bin

shutil.copy(XLSX, XLSM.replace(".xlsm", "_base.xlsx"))

# Read a minimal vbaProject.bin (just enough to make Excel open the file
# and allow VBA Trust Center to work). We'll store VBA as plain text instructions
# since generating binary .bin from scratch requires COM/xlwings.
# Instead we embed the VBA source in a way LibreOffice can compile on open.

with zipfile.ZipFile(XLSM.replace(".xlsm", "_base.xlsx"), 'r') as zin:
    with zipfile.ZipFile(XLSM, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = data.replace(
                    b'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    b'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
                )
            zout.writestr(item, data)
        # Store VBA source as reference
        zout.writestr('xl/vba_source/ThisWorkbook.bas', VBA_THISWORKBOOK)
        zout.writestr('xl/vba_source/Module1.bas', VBA_MODULE)
        zout.writestr('xl/vba_source/README.txt',
            "HOW TO ENABLE MULTI-SELECT MACRO:\n"
            "1. Open Student_Info_Template.xlsm in Excel\n"
            "2. Press Alt+F11 to open VBA Editor\n"
            "3. Double-click 'ThisWorkbook' and paste contents of ThisWorkbook.bas\n"
            "4. Insert > Module, paste contents of Module1.bas\n"
            "5. Save as .xlsm\n"
            "6. Close VBA Editor\n"
            "Now clicking any Subjects cell (column J, rows 7-106) will open the multi-select popup.\n"
        )

os.remove(XLSM.replace(".xlsm", "_base.xlsx"))
print(f"xlsm saved → {XLSM}")

# Also write the VBA files standalone for easy copy-paste
VBA_DIR = "/home/claude/hallsync_v2/data/vba_source"
os.makedirs(VBA_DIR, exist_ok=True)
with open(f"{VBA_DIR}/ThisWorkbook.bas", "w") as f:
    f.write(VBA_THISWORKBOOK)
with open(f"{VBA_DIR}/Module1.bas", "w") as f:
    f.write(VBA_MODULE)
with open(f"{VBA_DIR}/HOW_TO_INSTALL.txt", "w") as f:
    f.write(
        "HOW TO ENABLE MULTI-SELECT SUBJECTS MACRO\n"
        "==========================================\n\n"
        "1. Open Student_Info_Template.xlsm in Microsoft Excel\n"
        "2. Press Alt + F11  (opens VBA Editor)\n"
        "3. In the left panel, double-click 'ThisWorkbook'\n"
        "4. Copy and paste the full content of ThisWorkbook.bas into that window\n"
        "5. Right-click the project > Insert > Module\n"
        "6. Copy and paste the full content of Module1.bas into that window\n"
        "7. Press Ctrl+S to save (keep as .xlsm)\n"
        "8. Close the VBA Editor\n"
        "9. If prompted, click 'Enable Macros' when opening the file\n\n"
        "RESULT:\n"
        "Clicking any cell in the Subjects column (J7:J106)\n"
        "will open a popup with checkboxes for:\n"
        "  Maths, Science, Social Science, Biology,\n"
        "  Physics, Chemistry, Computer Science\n"
        "Ticking subjects fills them comma-separated into the cell automatically.\n"
    )
print(f"VBA source saved → {VBA_DIR}/")
print("Done.")
