# Portfolio Validator (Positional Scan)

A PySide6 desktop application for validating portfolio evidence against an Excel rubric and a structured evidence folder.

The tool scans:
- An Excel file (no headers, fixed column positions)
- A `Bewijslasten/` directory with subfolders per evidence ID (A–Z)

Each evidence item is visualized as a status card and can be inspected in detail.

---

## Features

- Visual status cards per evidence ID
- Status detection:
  - ✅ GOOD
  - 🟠 INCOMPLETE
  - ⚠️ EMPTY FOLDER
  - ❌ NOT RATED YET 
  - ❓ NO DATA
- Clickable evidence files (opens with default application)
- REA validation (Relevant, Authentiek, Actueel)
- Automatic TODO list per evidence item
- Refresh lockout (3 seconds) to prevent crashes
- Works cross-platform (Linux / Windows / macOS)

## Folder Structure

Your project folder must look like this:

```ProjectFolder/
├── portfolio.xlsx
└── Bewijslasten/
├── A/
│ ├── evidence1.pdf
│ └── screenshot.png
├── B/
└── C/
```

- The Excel file must be in the root
- Evidence folders must be named `A`–`Z`

## Installation

### 1. Python
Requires **Python 3.10+**

Check:
```bash
python --version
```

2. Create virtual environment (recommended)

```python -m venv .venv
source .venv/bin/activate        # Linux / macOS
.venv\Scripts\activate           # Windows
```
3. Install dependencies
```
pip install pandas openpyxl PySide6
```
Running the App
```
python Checker.py
```

Then:


1. Click Select Project Folder

2. Choose the folder containing the Excel file and Bewijslasten

3. Click cards to inspect evidence