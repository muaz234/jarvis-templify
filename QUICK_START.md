# QUICK START (Copy & Paste)

## First Time Setup

Open Terminal and copy-paste these commands one by one:

```bash
cd /Users/rhbdigital/Downloads/amy
python3 -m venv venv
source venv/bin/activate
pip install pandas openpyxl python-docx lxml pillow
```

## Every Time You Run

Open Terminal and copy-paste:

```bash
cd /Users/rhbdigital/Downloads/amy
source venv/bin/activate
python3 main.py
```

## What Happens

✅ The script reads `clients.xlsx`  
✅ Fills in template files with client data  
✅ Converts to PDF  
✅ Saves all PDFs to `output/` folder  

## If It Breaks

1. Check `generation_errors.log` file for error messages
2. Make sure template files exist in `templates/` folder
3. Make sure LibreOffice is installed
4. See `SETUP.md` for detailed troubleshooting

---

**Excel File Format Needed:**

| client_name | bank_statement | flight_ticket | hotel_booking | acc_number | booking_number | ticket_number | booking_reff |
|---|---|---|---|---|---|---|---|
| John Smith | BNI.docx | AA.docx | Hotel.docx | 123456 | ABC789 | T001 | REF123 |

- Template columns hold the **filename** of the Word template
- Other columns hold the **data** to fill into templates
