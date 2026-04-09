# Beginner Guide: Generate Client PDFs on macOS

This guide explains how to use the Python script `generate_documents_mac.py`.

The script does this:

1. reads client data from `clients.xlsx`
2. fills Word templates with that data
3. saves temporary `.docx` files
4. converts those `.docx` files into `.pdf`

You do **not** need to understand every line of Python to use it.

---

## 1. What you need first

You need these tools installed on your Mac:

- Python 3
- LibreOffice
- a few Python packages

### Install Python

Open **Terminal** and check if Python is already installed:

```bash
python3 --version
```

If you see a version like `Python 3.11.0`, you are good.

If not, install it with Homebrew:

```bash
brew install python
```

---

## 2. Install LibreOffice

This script uses **LibreOffice** to convert Word files to PDF.

Install it with Homebrew:

```bash
brew install --cask libreoffice
```

After that, check if LibreOffice is available:

```bash
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```

If it shows a version number, that part is ready.

---

## 3. Install Python packages

In Terminal, run:

```bash
pip3 install pandas openpyxl docxtpl
```

### What these packages do

- `pandas` reads Excel files
- `openpyxl` helps Python read `.xlsx`
- `docxtpl` fills Word templates with your data

---

## 4. Put your files in one folder

Create a folder, for example:

```bash
mkdir -p ~/client_pdf_project/templates
mkdir -p ~/client_pdf_project/output
mkdir -p ~/client_pdf_project/temp_docx
```

Now put these files inside `~/client_pdf_project`:

- `generate_documents_mac.py`
- `clients.xlsx`
- folder `templates/`

Inside `templates/`, put your Word templates:

- `letter.docx`
- `brochure.docx`
- `form.docx`

Your folder should look like this:

```text
client_pdf_project/
├── generate_documents_mac.py
├── clients.xlsx
├── templates/
│   ├── letter.docx
│   ├── brochure.docx
│   └── form.docx
├── output/
└── temp_docx/
```

---

## 5. Prepare your Excel file

Your Excel file **must** contain a column called:

```text
client_name
```

Example:

| client_name | address | phone | email |
|------------|---------|-------|-------|
| Ali Sdn Bhd | Kuala Lumpur | 0123456789 | ali@email.com |
| Maju Tech | Shah Alam | 0198887777 | maju@email.com |

Every column in Excel can be used inside the Word template.

So if your Excel has these columns:

- `client_name`
- `address`
- `phone`
- `email`

then your template can use these placeholders:

```text
{{ client_name }}
{{ address }}
{{ phone }}
{{ email }}
```

---

## 6. Prepare your Word templates

Open `letter.docx`, `brochure.docx`, or `form.docx` in Word.

Write normal text, then place variables using double curly brackets.

Example template content:

```text
Dear {{ client_name }},

Your address is {{ address }}.
Your phone number is {{ phone }}.
Your email is {{ email }}.
```

When the script runs, those placeholders will be replaced with data from Excel.

---

## 7. Run the script

Open Terminal and move into your project folder:

```bash
cd ~/client_pdf_project
```

Run the script like this:

```bash
python3 generate_documents_mac.py
```

---

## 8. What happens after running

The script will:

- read each row in `clients.xlsx`
- generate a `.docx` file for each template
- convert it to PDF
- save PDFs into the `output/` folder

Example output files:

```text
output/
├── Ali Sdn Bhd - Letter.pdf
├── Ali Sdn Bhd - Brochure.pdf
├── Ali Sdn Bhd - Form.pdf
├── Maju Tech - Letter.pdf
├── Maju Tech - Brochure.pdf
└── Maju Tech - Form.pdf
```

Temporary Word files are saved here:

```text
temp_docx/
```

---

## 9. If something goes wrong

The script creates a file called:

```text
generation_errors.log
```

That file contains the rows or templates that failed.

### Common problems

#### Problem: `Client file not found`
Your `clients.xlsx` file is missing or not in the same folder as the script.

#### Problem: `Excel must contain a 'client_name' column`
Your Excel file does not have a column named exactly `client_name`.

#### Problem: `Missing template`
One of these files is missing:

- `templates/letter.docx`
- `templates/brochure.docx`
- `templates/form.docx`

#### Problem: `LibreOffice was not found`
LibreOffice is not installed yet, or the script cannot find it.

Install it with:

```bash
brew install --cask libreoffice
```

---

## 10. How the script works in simple words

Here is the flow:

1. make sure folders exist
2. open Excel file
3. read one client row
4. fill one Word template
5. save temporary `.docx`
6. convert `.docx` to `.pdf`
7. repeat for all clients and all templates

---

## 11. How to test with only a few rows

Inside the script, you will see this line:

```python
TEST_LIMIT = None
```

You can change it to this:

```python
TEST_LIMIT = 2
```

That means only the first 2 clients in Excel will be processed.

This is useful while testing.

---

## 12. What changed from the Windows version

The original version used Microsoft Word COM automation, which works on Windows.

Your Mac version is different:

- removed `win32com.client`
- removed `pywin32`
- added LibreOffice conversion using `subprocess`

So this version is more suitable for macOS.

---

## 13. Beginner tips

- Start with just 1 client row in Excel
- Start with just 1 simple template
- Make sure your placeholder names match your Excel column names exactly
- Keep file names simple
- Check `generation_errors.log` if anything fails

---

## 14. Quick start checklist

Before running, make sure all of this is true:

- Python installed
- LibreOffice installed
- `pandas`, `openpyxl`, `docxtpl` installed
- `clients.xlsx` exists
- Excel has `client_name` column
- `templates/letter.docx` exists
- `templates/brochure.docx` exists
- `templates/form.docx` exists

---

## 15. One full example

### Excel columns

```text
client_name, address, phone
```

### One row of data

```text
Ali Sdn Bhd, Kuala Lumpur, 0123456789
```

### Template text

```text
Hello {{ client_name }}
Address: {{ address }}
Phone: {{ phone }}
```

### Result PDF text

```text
Hello Ali Sdn Bhd
Address: Kuala Lumpur
Phone: 0123456789
```

---

## 16. Command summary

Install Python packages:

```bash
pip3 install pandas openpyxl docxtpl
```

Install LibreOffice:

```bash
brew install --cask libreoffice
```

Go to the project folder:

```bash
cd ~/client_pdf_project
```

Run the script:

```bash
python3 generate_documents_mac.py
```

---

## 17. Final note

You are basically doing a simple mail-merge system:

- Excel = your data
- Word templates = your document design
- Python = the automation
- LibreOffice = the PDF converter

Once your folder structure and template names are correct, the script should be easy to run again and again.
