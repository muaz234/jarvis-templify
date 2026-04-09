"""
DOCUMENT GENERATOR
==================

This script automatically generates PDF files from Word templates and client data.
Compatible with macOS and Windows.

HOW TO RUN (macOS):
-------------------
1. Open Terminal
2. cd /Users/YOUR_USERNAME/Downloads/project-dir
3. source venv/bin/activate
4. python3 main.py

HOW TO RUN (Windows):
---------------------
1. Open Command Prompt
2. cd C:\Users\YOUR_USERNAME\Downloads\project-dir
3. venv\Scripts\activate
4. python main.py

WHAT YOU NEED:
--------------
- clients.xlsx (Excel file with client data)
- templates/ folder (with your Word template files)
- LibreOffice installed (for converting DOCX to PDF)
- Python 3.7+ with pandas and docxtpl installed

For detailed setup instructions, see: OPERATING.md
"""

from __future__ import annotations

import os
import re
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Any, List, Optional

import pandas as pd
from docxtpl import DocxTemplate


# =========================
# CONFIG
# =========================
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
TEMP_DOCX_DIR = BASE_DIR / "temp_docx"
CLIENT_FILE = BASE_DIR / "clients.xlsx"

# The three template columns - if they have a value, that's the template filename to use
TEMPLATE_COLUMNS = ["bank_statement", "flight_ticket", "hotel_booking"]

# Optional: only process the first N clients for testing
# Example: TEST_LIMIT = 5
TEST_LIMIT: Optional[int] = None


# =========================
# HELPERS
# =========================
def sanitize_filename(name: str) -> str:
    """
    Make a string safe to use as a filename.
    """
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180] if name else "Unnamed"


def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEMP_DOCX_DIR.mkdir(parents=True, exist_ok=True)


def load_clients(xlsx_path: Path) -> pd.DataFrame:
    """
    Read the Excel file containing client data.
    """
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Client file not found: {xlsx_path}")

    df = pd.read_excel(xlsx_path)

    if "client_name" not in df.columns:
        raise ValueError("Excel must contain a 'client_name' column.")

    # Replace empty Excel cells so templates receive empty strings instead of NaN
    df = df.fillna("")

    if TEST_LIMIT:
        df = df.head(TEST_LIMIT)

    return df


def build_context(row: pd.Series) -> Dict[str, Any]:
    """
    Turn one Excel row into a dictionary used by the Word template.
    Every Excel column becomes available inside the template.
    """
    context: Dict[str, Any] = {}

    for col in row.index:
        value = row[col]
        if pd.isna(value):
            value = ""
        context[col] = str(value)

    return context


def render_docx(template_path: Path, context: Dict[str, Any], output_docx_path: Path) -> None:
    """
    Fill a .docx template with data and save the rendered document.
    """
    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(output_docx_path))


def find_soffice() -> Optional[str]:
    """
    Try to find LibreOffice's 'soffice' command on macOS or Windows.
    """
    candidates = []
    
    # Check PATH first (works on all systems)
    soffice_in_path = shutil.which("soffice")
    if soffice_in_path:
        candidates.append(soffice_in_path)
    
    # macOS paths
    candidates.extend([
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
    ])
    
    # Windows paths
    candidates.extend([
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ])
    
    for candidate in candidates:
        try:
            if candidate and Path(candidate).exists():
                return str(candidate)
        except (OSError, ValueError):
            # Handle invalid paths that Path() might reject
            continue
    
    return None


def export_docx_to_pdf(input_docx: Path, output_pdf: Path) -> None:
    """
    Convert DOCX to PDF using LibreOffice in headless mode.

    Works on both macOS and Windows if LibreOffice is installed.
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice was not found. Please install LibreOffice first, "
            "or make sure the 'soffice' command is available."
        )

    # LibreOffice writes the PDF into the directory given by --outdir.
    # It keeps the same base filename as the input DOCX.
    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_pdf.parent),
            str(input_docx),
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    expected_pdf = output_pdf.parent / f"{input_docx.stem}.pdf"

    if not expected_pdf.exists():
        raise RuntimeError(
            f"LibreOffice finished, but the PDF was not found: {expected_pdf}"
        )

    if expected_pdf != output_pdf:
        expected_pdf.replace(output_pdf)


# =========================
# MAIN PROGRAM
# =========================
def main() -> int:
    ensure_dirs()

    try:
        df = load_clients(CLIENT_FILE)
    except Exception as exc:
        print(f"[ERROR] Failed to load client list: {exc}")
        return 1

    success_count = 0
    failures: List[str] = []

    for idx, row in df.iterrows():
        client_name = sanitize_filename(row.get("client_name", "Unnamed"))
        context = build_context(row)

        # Check each template column (bank_statement, flight_ticket, hotel_booking)
        for template_column in TEMPLATE_COLUMNS:
            template_filename = row.get(template_column, "").strip()
            
            # Skip if this template column is empty for this client
            if not template_filename:
                continue

            # Check if template file exists
            template_path = TEMPLATE_DIR / template_filename

            if not template_path.exists():
                failures.append(
                    f"Row {idx + 2} | {client_name} | {template_column} | Template file not found: {template_filename}"
                )
                continue

            # Create output filename with template column name
            pdf_name = f"{client_name} - {template_column}.pdf"
            temp_docx_name = f"{client_name} - {template_column}.docx"

            temp_docx_path = TEMP_DOCX_DIR / temp_docx_name
            pdf_path = OUTPUT_DIR / pdf_name

            try:
                render_docx(template_path, context, temp_docx_path)
                export_docx_to_pdf(temp_docx_path, pdf_path)
                success_count += 1
                print(f"[OK] {pdf_path.name}")
            except subprocess.CalledProcessError as exc:
                failures.append(
                    f"Row {idx + 2} | {client_name} | {template_column} | "
                    f"LibreOffice conversion error: {exc.stderr.strip() or exc}"
                )
            except Exception as exc:
                failures.append(
                    f"Row {idx + 2} | {client_name} | {template_column} | Error: {exc}"
                )

    print("\n====================")
    print(f"PDFs created: {success_count}")
    print(f"Failures: {len(failures)}")

    if failures:
        log_path = BASE_DIR / "generation_errors.log"
        with open(log_path, "w", encoding="utf-8") as file:
            for item in failures:
                file.write(item + "\n")
        print(f"Error log saved to: {log_path}")

    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())
