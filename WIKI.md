# Document Generator - User Guide

Welcome! This wiki explains how to use the Document Generator in simple terms.

---

## What Is This?

The **Document Generator** is a tool that automatically creates PDF documents from templates using information from an Excel spreadsheet.

**Simple analogy:** Imagine you have a template letter that you want to send to 100 customers. Instead of manually changing the names and account numbers for each one, this tool does it automatically in seconds.

### What Can You Generate?

- 📄 Bank Statements
- ✈️ Flight Tickets
- 🏨 Hotel Confirmations
- 📋 Any custom document with template variables

---

## How It Works

### Step 1: You Prepare
- Create an Excel file with your customer/client information
- Prepare Word template files with placeholder text

### Step 2: Tool Processes
- Reads your Excel data
- Fills in the template files with each client's information
- Converts to PDF

### Step 3: You Get
- Professional PDF files ready to use
- All PDFs saved in the output folder
- Error log if anything goes wrong

```
┌─────────────────┐
│   clients.xlsx  │
│  (Your data)    │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Word Templates  │
│ (letter.docx)   │
└────────┬────────┘
         │
         ▼
    ┌────────────────────┐
    │ Document Generator │
    │    (main.py)       │
    └────────┬───────────┘
             │
             ▼
    ┌────────────────────┐
    │    PDF Files       │
    │  (output folder)   │
    └────────────────────┘
```

---

## Getting Started

### Before You Start

Make sure you have:
- ✅ A Mac or Windows computer
- ✅ About 15-20 minutes for first-time setup
- ✅ Your Excel file with client data
- ✅ Your Word template files

### Setup (First Time Only)

**For Mac:**
1. Open Terminal (search "Terminal" in Spotlight)
2. Copy and paste these commands:

```bash
cd /Users/YOUR_USERNAME/Downloads/project-dir
python3 -m venv venv
source venv/bin/activate
pip install pandas openpyxl python-docx lxml pillow
```

**For Windows:**
1. Open Command Prompt (search "Command Prompt" or press Win+R and type cmd)
2. Copy and paste these commands:

```bash
cd C:\Users\YOUR_USERNAME\Downloads\project-dir
python -m venv venv
venv\Scripts\activate
pip install pandas openpyxl python-docx lxml pillow
```

Replace `YOUR_USERNAME` with your actual username.

**What to expect:** You'll see text appearing as the computer installs software. This takes 2-3 minutes. When done, you should see "Successfully installed..."

---

## Preparing Your Excel File

Your Excel file needs specific columns:

### Required Column
- **client_name** - The name of each person/company

### Template Selection Columns
- **bank_statement** - Put the template filename here (e.g., "BNI_Statement.docx")
- **flight_ticket** - Put the template filename here (e.g., "AA_Flight.docx")
- **hotel_booking** - Put the template filename here (e.g., "Agoda_Hotel.docx")

### Data Columns (used to fill templates)
- **acc_number** - Account number for bank statements
- **booking_number** - For flight reservations
- **ticket_number** - For flight tickets
- **booking_reff** - For hotel bookings

### Example Excel File

| client_name | bank_statement | acc_number | flight_ticket | booking_number | ticket_number | hotel_booking | booking_reff |
|---|---|---|---|---|---|---|---|
| John Smith | BNI_Statement.docx | 1234567890 | AA_Flight.docx | ABC123 | T12345 | Agoda_Hotel.docx | REF001 |
| Sarah Johnson | (leave empty) | | AA_Flight.docx | XYZ789 | T54321 | (leave empty) | |
| Mike Brown | CIMB_Statement.docx | 9876543210 | (leave empty) | | | Booking_Hotel.docx | REF002 |

**Notes:**
- Leave cells empty if you don't want to generate that type of document
- Template filenames must match exactly (including .docx)
- All data fields must be filled for templates you want to use

---

## Preparing Your Templates

Templates are Word documents (.docx) with placeholder text that gets replaced.

### Example Template (Word document)

```
BANK STATEMENT

Account Holder: {{ client_name }}
Account Number: {{ acc_number }}
Date: [Today's Date]

Your account is active.
```

When the tool runs with John Smith's data:
```
BANK STATEMENT

Account Holder: John Smith
Account Number: 1234567890
Date: [Today's Date]

Your account is active.
```

### How to Create Placeholders

1. Open your template in Microsoft Word or LibreOffice
2. Where you want data to go, type: `{{ column_name }}`
3. Examples:
   - `{{ client_name }}`
   - `{{ acc_number }}`
   - `{{ booking_number }}`
   - `{{ ticket_number }}`
   - `{{ booking_reff }}`

4. Save the file as .docx in your `templates/` folder

---

## Running the Tool

### Every Time You Use It

**For Mac:**
```bash
cd /Users/YOUR_USERNAME/Downloads/project-dir
source venv/bin/activate
python3 main.py
```

**For Windows:**
```bash
cd C:\Users\YOUR_USERNAME\Downloads\project-dir
venv\Scripts\activate
python main.py
```

### What Happens

The tool will show messages like:
```
[OK] John Smith - bank_statement.pdf
[OK] John Smith - flight_ticket.pdf
[OK] Sarah Johnson - flight_ticket.pdf
...
====================
PDFs created: 12
Failures: 0
```

Check your `output/` folder - your PDFs are there!

---

## File Organization

Your folder should look like this:

```
project-dir/
├── main.py                          ← The script ✅
├── clients.xlsx                     ← Your data ✅
├── templates/                       ← Your Word files ✅
│   ├── BNI_Statement.docx
│   ├── AA_Flight.docx
│   └── Agoda_Hotel.docx
├── output/                          ← PDFs appear here (auto-created)
├── temp_docx/                       ← Temporary files (auto-created)
├── README.md                        ← Info file ✅
├── OPERATING.md                     ← Setup instructions ✅
└── .gitignore                       ← Git config ✅
```

---

## FAQ (Frequently Asked Questions)

### Q: Can I leave some cells empty?
**A:** Yes! If you don't have a bank statement for someone, just leave that cell empty. The tool will skip it for that client.

### Q: What if I only want to generate flight tickets?
**A:** Just leave the other template columns empty. Only fill in flight_ticket with your template filename.

### Q: How many documents can I create at once?
**A:** As many as you want! Hundreds, thousands - the tool will process them all.

### Q: Can I update templates later?
**A:** Yes! Exit the tool, update your template or Excel file, and run it again.

### Q: Where are my PDFs?
**A:** Look in the `output/` folder. You'll see files like "John Smith - bank_statement.pdf"

### Q: What if something goes wrong?
**A:** Check the `generation_errors.log` file in your main folder. It explains what went wrong.

---

## Troubleshooting

### Problem: "Python is not found" or "command not found"

**Solution (Mac):**
- Python isn't installed
- Download from: https://www.python.org/downloads/
- Install and restart your computer

**Solution (Windows):**
- Python wasn't added to PATH during installation
- Uninstall Python from Control Panel
- Reinstall and **check the box "Add Python to PATH"**

### Problem: "LibreOffice was not found"

**Solution:**
- LibreOffice isn't installed
- Download from: https://www.libreoffice.org/download/
- Install and restart your computer

### Problem: No PDFs are created but no error message

**Solution:**
- Your template column cells might be empty
- Check your Excel file - make sure you have template filenames in the bank_statement/flight_ticket/hotel_booking columns
- Make sure template files actually exist in the templates/ folder

### Problem: "Template file not found"

**Solution:**
- Check the filename spelling carefully (case sensitive)
- Make sure the file is in the `templates/` folder
- File name must match exactly what's in your Excel file

### Problem: "Excel must contain a 'client_name' column"

**Solution:**
- Your Excel file needs a column named exactly "client_name"
- Add this column with your client names

### Problem: Can't activate virtual environment

**Mac:**
```bash
source venv/bin/activate
```

**Windows:**
```bash
venv\Scripts\activate
```

If that doesn't work, you might need to recreate the virtual environment:

**Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

---

## Tips & Tricks

### Tip 1: Test with a Small Batch First
Before generating 1000 documents, try with 5 clients first to make sure everything works.

### Tip 2: Double-Check Template Names
Type template names (e.g., "BNI_Statement.docx") exactly as they appear in your files, including case and spacing.

### Tip 3: Keep Template Files Simple
Complex template formatting might not convert perfectly to PDF. Test with a sample first.

### Tip 4: Save Your Templates
Always keep backup copies of your template files before running the tool.

### Tip 5: Schedule Regular Runs
If you generate documents regularly, set a reminder to run the tool at a specific time.

---

## Common Tasks

### How to Edit the Excel File
1. Open Excel
2. Load your `clients.xlsx` file
3. Edit as needed
4. Save (Ctrl+S or Cmd+S)
5. Run the tool again

### How to Add a New Template Type
1. Create your new Word template with placeholders
2. Save as .docx in the `templates/` folder
3. You'll need to edit `main.py` to add support for the new type
4. (For this, contact your technical person)

### How to Keep Old PDFs
PDFs are created with dates in filenames, so each time you run the tool, new files are created. Old ones are not deleted automatically.

### How to Share Results
All your PDFs are in the `output/` folder. You can:
- Email them individually
- Zip the entire folder and share
- Upload to cloud storage (Google Drive, Dropbox, etc.)

---

## Next Steps

1. ✅ Install Python and LibreOffice
2. ✅ Run the setup commands
3. ✅ Create your Excel file
4. ✅ Create your Word templates
5. ✅ Run the tool
6. ✅ Check the output folder for PDFs

**You're done!**

---

## Need Help?

- Check `OPERATING.md` for detailed setup instructions
- Look in `generation_errors.log` for error details
- Review the template examples above
- Refer to the FAQ section

**Still stuck?** Make sure:
- ✅ Python is installed and working
- ✅ LibreOffice is installed
- ✅ Virtual environment is activated (you see `(venv)` in terminal)
- ✅ Excel file columns match exactly
- ✅ Template filenames match exactly
- ✅ All template files are in the `templates/` folder
