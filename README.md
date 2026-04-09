# Document Generator

Automatically generate PDF documents from templates and client data.

## What It Does

This tool reads client information from an Excel spreadsheet and creates professional PDF documents by filling in Word templates with the client data. Perfect for generating bank statements, flight tickets, hotel confirmations, and more in bulk.

## Key Features

✅ Batch processing - generate hundreds of PDFs at once  
✅ Multiple templates per client - create different documents for each client  
✅ Flexible data mapping - use any Excel column names  
✅ Error tracking - saves all errors to a log file  
✅ Cross-platform - works on Mac and Windows  

## Requirements

- Excel file with client data (`clients.xlsx`)
- Word template files with placeholder text
- Python 3.7 or higher
- LibreOffice (free) for PDF conversion
- 5-10 minutes to set up

## Quick Start

1. **Set up the environment** (first time only):
   - Follow the instructions in `OPERATING.md`

2. **Prepare your data**:
   - Create `clients.xlsx` with your client information
   - Add Word templates to the `templates/` folder

3. **Run the generator**:
   ```bash
   python3 main.py
   ```

4. **Get your PDFs**:
   - Check the `output/` folder for your generated files

## File Structure

```
project-folder/
├── main.py                 # The script (run this)
├── clients.xlsx            # Your client data (Excel file)
├── templates/              # Your Word templates (create this folder)
│   ├── BNI_Statement.docx
│   ├── AA_Flight.docx
│   └── Agoda_Hotel.docx
├── output/                 # Generated PDFs (created automatically)
├── temp_docx/              # Temporary files (created automatically)
└── OPERATING.md            # Setup instructions
```

## Excel Format

Your Excel file should have:
- A `client_name` column
- Columns for each template type (e.g., `bank_statement`, `flight_ticket`, `hotel_booking`)
- Data columns (e.g., `acc_number`, `booking_number`, `ticket_number`, `booking_reff`)

Example:

| client_name | bank_statement | acc_number | flight_ticket | booking_number | ticket_number | hotel_booking | booking_reff |
|---|---|---|---|---|---|---|---|
| John Smith | BNI.docx | 123456789 | AA.docx | ABC123 | T001 | Hotel.docx | REF123 |

## Need Help?

- See `OPERATING.md` for setup instructions (Mac & Windows)
- See `QUICK_START.md` for copy-paste commands
- Check `generation_errors.log` if something goes wrong

## License

Free to use and modify.
