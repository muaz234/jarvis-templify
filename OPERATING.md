# Operating Instructions

Complete setup and running instructions for Mac and Windows.

---

## Prerequisites

### What You Need to Install

1. **Python 3.7+** (free)
2. **LibreOffice** (free)
3. **Text Editor** (optional - to edit `clients.xlsx`)

---

## macOS Setup

### Step 1: Install Python 3

If you don't have Python installed:

1. Visit https://www.python.org/downloads/
2. Download the **macOS installer**
3. Run the installer and follow the prompts
4. Restart your Mac

**Verify installation:**
```bash
python3 --version
```

Should show: `Python 3.X.X`

### Step 2: Install LibreOffice

1. Visit https://www.libreoffice.org/download/download-libreoffice/
2. Download macOS version
3. Double-click the DMG file
4. Drag LibreOffice to Applications folder
5. Wait for installation to complete

### Step 3: Create Virtual Environment (First Time Only)

Open Terminal and run:

```bash
cd /Users/rhbdigital/Downloads/amy
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install pandas openpyxl python-docx lxml pillow
```

This creates a virtual environment and installs dependencies. It may take 2-3 minutes.

**You should see:**
```
Successfully installed pandas openpyxl python-docx lxml pillow
```

### Step 4: Prepare Your Files

1. Create a `templates/` folder in your project directory
2. Place your Word template files inside (`BNI.docx`, `AA.docx`, etc.)
3. Create `clients.xlsx` with your client data

### Step 5: Run the Script

Every time you want to generate PDFs:

```bash
cd /Users/rhbdigital/Downloads/amy
source venv/bin/activate
python3 main.py
```

**Expected output:**
```
[OK] John Smith - bank_statement.pdf
[OK] John Smith - flight_ticket.pdf
[OK] John Smith - hotel_booking.pdf
...
====================
PDFs created: 12
Failures: 0
```

### Step 6: Exit Virtual Environment

When done:
```bash
deactivate
```

---

## Windows Setup

### Step 1: Install Python 3

1. Visit https://www.python.org/downloads/
2. Click **Download Python 3.X.X** (the big yellow button)
3. Run the installer
4. ⚠️ **IMPORTANT**: Check the box "Add Python to PATH"
5. Click "Install Now"
6. Wait for installation to complete
7. Restart your computer

**Verify installation:**
Open Command Prompt and type:
```bash
python --version
```

Should show: `Python 3.X.X`

### Step 2: Install LibreOffice

1. Visit https://www.libreoffice.org/download/download-libreoffice/
2. Download Windows version
3. Run the installer
4. Follow the installation wizard
5. Wait for installation to complete

### Step 3: Create Virtual Environment (First Time Only)

Open Command Prompt:

1. Press `Win + R`
2. Type `cmd` and press Enter

Now paste these commands one by one:

```bash
cd C:\Users\YourUserName\Downloads\amy
python -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
pip install pandas openpyxl python-docx lxml pillow
```

Replace `YourUserName` with your actual Windows username.

**You should see:**
```
Successfully installed pandas openpyxl python-docx lxml pillow
```

### Step 4: Prepare Your Files

1. Create a `templates/` folder in your project directory
2. Place your Word template files inside
3. Create `clients.xlsx` with your client data

### Step 5: Run the Script

Every time you want to generate PDFs:

Open Command Prompt:

```bash
cd C:\Users\YourUserName\Downloads\amy
venv\Scripts\activate
python main.py
```

**Expected output:**
```
[OK] John Smith - bank_statement.pdf
[OK] John Smith - flight_ticket.pdf
[OK] John Smith - hotel_booking.pdf
...
====================
PDFs created: 12
Failures: 0
```

### Step 6: Exit Virtual Environment

When done:
```bash
deactivate
```

---

## Dependencies Summary

| Dependency | Purpose | Installation |
|---|---|---|
| Python 3.7+ | Programming language | https://www.python.org/downloads/ |
| pandas | Reading Excel files | `pip install pandas` |
| openpyxl | Excel support | `pip install openpyxl` |
| python-docx | Creating Word documents | `pip install python-docx` |
| lxml | XML processing | `pip install lxml` |
| pillow | Image processing | `pip install pillow` |
| LibreOffice | Converting DOCX to PDF | https://www.libreoffice.org/download/ |

---

## Troubleshooting

### "python: command not found" (Mac/Linux) or "python is not recognized" (Windows)

**Mac:**
- Python isn't installed. Install from https://www.python.org/downloads/

**Windows:**
- Python wasn't added to PATH. 
- Uninstall Python, reinstall, and **check "Add Python to PATH"** during installation

### "LibreOffice was not found"

- LibreOffice isn't installed or not in the expected location
- Install from https://www.libreoffice.org/download/
- Restart your computer after installation

### "Excel must contain a 'client_name' column"

- Your Excel file doesn't have a `client_name` column
- Add this column to your Excel sheet

### "Template file not found: filename.docx"

- The template file name doesn't match what's in Excel
- Check spelling and make sure file is in `templates/` folder

### Script runs but generates no PDFs

- Check your Excel file has the correct column names
- Verify template files exist in `templates/` folder
- Check `generation_errors.log` for detailed errors

### "No module named 'pandas'" or similar

- Virtual environment not activated
- Make sure you see `(venv)` at the start of the command line
- If not, activate it: `source venv/bin/activate` (Mac) or `venv\Scripts\activate` (Windows)

---

## Common Commands

### Mac

```bash
# Navigate to project
cd /Users/USERNAME/Downloads/amy

# Activate virtual environment
source venv/bin/activate

# Run script
python3 main.py

# Deactivate
deactivate
```

### Windows

```bash
# Navigate to project
cd C:\Users\USERNAME\Downloads\amy

# Activate virtual environment
venv\Scripts\activate

# Run script
python main.py

# Deactivate
deactivate
```

---

## Need More Help?

1. Check the error message in `generation_errors.log`
2. Review the Excel file format in `README.md`
3. Verify all template files exist
4. Make sure LibreOffice is installed and working
