# HallSync — First Time Setup Guide

Follow these steps exactly the first time you set up HallSync on your computer.

---

## Step 1 — Install Python

Download and install Python 3.10 or newer from https://python.org/downloads

During installation, check the box that says **"Add Python to PATH"**.

Verify it worked by opening Terminal / Command Prompt and typing:
```
python --version
```
You should see something like `Python 3.12.0`.

---

## Step 2 — Download HallSync

Go to the GitHub page and click the green **Code** button → **Download ZIP**.

Extract the ZIP to a folder on your computer, for example:
```
C:\Users\YourName\HallSync\
```

---

## Step 3 — Install required libraries

Open Terminal / Command Prompt, navigate to the HallSync folder:
```
cd C:\Users\YourName\HallSync
```

Then run:
```
pip install -r requirements.txt
```

This installs pandas, openpyxl, and reportlab automatically.

---

## Step 4 — Generate the Excel templates (one time only)

```
python scripts/step1_create_student_template.py
python scripts/step2_create_school_template.py
```

This creates:
- `data/Student_Info_Template.xlsm`  — give this to each school
- `data/School_Info_Template.xlsx`   — fill this yourself as admin

---

## Step 5 — Enable the Subjects macro in the student template

1. Open `data/Student_Info_Template.xlsm` in **Microsoft Excel**
2. Press **Alt + F11** to open the VBA Editor
3. Double-click **ThisWorkbook** in the left panel
4. Copy and paste the contents of `vba/ThisWorkbook.bas` into that window
5. Right-click the project → **Insert → Module**
6. Copy and paste the contents of `vba/Module1.bas` into that window
7. Press **Ctrl+S**, close the VBA Editor
8. When opening the file again, click **Enable Macros**

Now clicking any Subjects cell will open a multi-select popup.

---

## Step 6 — Fill data and run

1. Give `Student_Info_Template.xlsm` to each school to fill
2. Fill `School_Info_Template.xlsx` yourself
3. Drop all filled school files into `data/inbox/`
4. Run:
```
python main.py
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `pip` not found | Reinstall Python with "Add to PATH" checked |
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` again |
| Macros don't work | Use Microsoft Excel, not Google Sheets or LibreOffice |
| Validation errors | Read the error report and fix the Excel file, then re-run |
| Empty inbox | Put school Excel files in `data/inbox/` folder |
