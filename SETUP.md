# How to Run the Document Generator

This guide explains how to set up and run the document generator in simple terms.

## What Does This Do?

This script reads client information from an Excel file and automatically generates PDF documents from templates. For example:
- Bank statements with account numbers
- Flight tickets with booking numbers  
- Hotel confirmations with booking references

All it needs is your Excel file with the template filenames and data, and it creates professional PDFs automatically.

---

## Step-by-Step Setup (macOS)

### Step 1: Install Python 3 (if not already installed)

1. Download Python 3 from: https://www.python.org/downloads/
2. Run the installer and follow the prompts
3. Restart your computer

**Check if Python is installed:**
Open Terminal and type:
```bash
python3 --version
```
You should see something like: `Python 3.11.0`

---

### Step 2: Create a Virtual Environment

A virtual environment is like a separate workspace for this project (keeps things organized).

1. Open Terminal
2. Go to the project folder:
```bash
cd /Users/rhbdigital/Downloads/amy
```

3. Create the virtual environment:
```bash
python3 -m venv venv
```

This creates a folder called `venv` - don't delete it!

---

### Step 3: Activate the Virtual Environment

Every time you want to run the script, activate the environment first:

```bash
source venv/bin/activate
```

You'll see `(venv)` at the start of your terminal line - that means it's active.

---

### Step 4: Install Required Libraries

These are extra tools the script needs to work:

```bash
pip install pandas openpyxl python-docx lxml pillow
```

This may take a minute. Wait for it to finish.

---

### Step 5: Prepare Your Files

Make sure you have:

1. **`clients.xlsx`** — Your Excel file with client data
   - Must have a `client_name` column
   - Template columns: `bank_statement`, `flight_ticket`, `hotel_booking`
   - Data columns: `acc_number`, `booking_number`, `ticket_number`, `booking_reff`

2. **`templates/` folder** — Create this folder and put your Word templates inside
   - Example: `BNI_Statement.docx`, `AA_Flight.docx`, `Agoda_Hotel.docx`

3. Make sure **LibreOffice** is installed (for PDF conversion)
   - Download from: https://www.libreoffice.org/download/

Your folder structure should look like:
```
amy/
├── main.py
├── clients.xlsx
├── templates/
│   ├── BNI_Statement.docx
│   ├── AA_Flight.docx
│   └── Agoda_Hotel.docx
├── output/  (created automatically)
└── temp_docx/  (created automatically)
```

---

### Step 6: Run the Script

With the virtual environment active (see Step 3), type:

```bash
python3 main.py
```

The script will:
- Read your Excel file
- Generate PDFs in the `output/` folder
- Show success messages like `[OK] NOVIDA HANDAYANI - bank_statement.pdf`
- If there are errors, save them to `generation_errors.log`

---

## Quick Reference

**Every time you use this script:**

1. Open Terminal
2. Go to the folder: `cd /Users/rhbdigital/Downloads/amy`
3. Activate the environment: `source venv/bin/activate`
4. Run the script: `python3 main.py`
5. Check the `output/` folder for your PDFs

**To exit the virtual environment:**
```bash
deactivate
```

---

## Troubleshooting

**"python3: command not found"**
- Python isn't installed. Download it from https://www.python.org/downloads/

**"LibreOffice was not found"**
- Install LibreOffice from https://www.libreoffice.org/download/

**"Excel must contain a 'client_name' column"**
- Check your Excel file has a column named exactly `client_name`

**"Template file not found"**
- Check the template filename in your Excel matches the file in the `templates/` folder

**Any other error?**
- Check the `generation_errors.log` file for details

---

## Need Help?

If something isn't working:
1. Check the error log file (`generation_errors.log`)
2. Verify your Excel file structure matches the requirements
3. Make sure all template files exist in the `templates/` folder
