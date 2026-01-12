#!/usr/bin/env python3
"""
FNB Recipients to Investec CSV Converter

This script parses FNB recipient PDF exports and converts them to Investec's
beneficiary import CSV format.

Copyright (C) 2025 Ashley Kleynhans

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
"""

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Optional

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pytesseract
    from PIL import Image
    import io
except ImportError:
    pytesseract = None
    Image = None


# FNB uses universal branch code for electronic payments
FNB_BRANCH_CODE = "250655"

# Investec CSV headers
INVESTEC_HEADERS = [
    "Beneficiary Account Name",
    "Beneficiary Bank",
    "Beneficiary Bank Account Number",
    "Beneficiary Branch Code",
    "Beneficiary Reference",
    "Statement Description",
    "Beneficiary Name",
    "Beneficiary Fax Number",
    "Beneficiary Email Address",
    "Beneficiary Cell Number",
]


def extract_with_pdfplumber(pdf_path: str) -> list[dict]:
    """Extract recipient data using pdfplumber with word-level extraction."""
    if pdfplumber is None:
        raise ImportError("pdfplumber is not installed")

    recipients = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Use word-level extraction for better control
            words = page.extract_words(keep_blank_chars=False, x_tolerance=3, y_tolerance=3)
            page_recipients = parse_pdfplumber_words(words)
            recipients.extend(page_recipients)

    return recipients


def parse_pdfplumber_words(words: list[dict]) -> list[dict]:
    """Parse words from pdfplumber to extract recipient data."""
    recipients = []
    account_pattern = re.compile(r'^(\d{8,11})$')

    # Sort words by y position first, then x position
    words_sorted = sorted(words, key=lambda w: (w['top'], w['x0']))

    # Find account numbers and their positions
    account_positions = []
    for word in words_sorted:
        text = word['text'].strip()
        if account_pattern.match(text):
            account_positions.append({
                'account': text,
                'x': word['x0'],
                'y': word['top'],
            })

    # For each account number, find the associated name
    # Names are ABOVE the account number in FNB PDFs (within ~30 pixels above)
    for acc_info in account_positions:
        account = acc_info['account']
        acc_y = acc_info['y']
        acc_x = acc_info['x']

        # Find words that are above the account number (within 35 pixels)
        # and in the left column (name column starts around x=37)
        name_words = []
        for word in words_sorted:
            word_y = word['top']
            word_x = word['x0']
            text = word['text'].strip()

            # Skip the account number itself
            if text == account:
                continue

            # Check if word is above the account (within 35 pixels)
            # and at similar x position (left column, x < 130)
            y_diff = acc_y - word_y
            if 0 < y_diff <= 35 and word_x < 130:
                # Skip monetary amounts and dates
                if re.match(r'^[\d,]+\.\d{2}$', text.replace(' ', '')):
                    continue
                if re.match(r'^\d{2}$', text):  # Day part of date
                    continue
                if text in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                    continue
                if re.match(r'^20\d{2}$', text):  # Year
                    continue
                if text in ['0.00', 'Inactive', 'Recipient']:
                    continue

                name_words.append({'text': text, 'x': word_x, 'y': word_y})

        # Sort name words by y position (top to bottom), then x (left to right)
        if name_words:
            name_words.sort(key=lambda w: (w['y'], w['x']))
            name = ' '.join(w['text'] for w in name_words)
            name = clean_name(name)

            if name and len(name) > 1:
                recipients.append({
                    'name': name,
                    'account': account,
                    'reference': name
                })

    return recipients


def parse_fnb_table(table: list[list]) -> list[dict]:
    """Parse FNB recipient data from extracted table."""
    recipients = []
    account_pattern = re.compile(r'\b(\d{8,11})\b')

    for row in table:
        if not row or all(cell is None or cell == '' for cell in row):
            continue

        # Convert row to strings and clean
        row_data = [str(cell).strip() if cell else '' for cell in row]
        row_text = ' '.join(row_data)

        # Skip header rows and section headers
        skip_keywords = [
            'Name', 'Pay Amount', 'Last Paid', 'Amount', 'Their Reference',
            'My Reference', 'Please note', 'Due to system', 'Real-time'
        ]
        if any(keyword in row_text for keyword in skip_keywords):
            continue

        # Look for account number in row
        account_match = account_pattern.search(row_text)
        if account_match:
            account = account_match.group(1)

            # Try to find the name - usually first non-empty, non-numeric cell
            name = None
            reference = None

            for cell in row_data:
                if not cell:
                    continue
                # Skip if it's just the account number
                if cell == account:
                    continue
                # Skip monetary amounts
                if re.match(r'^[\d,]+\.\d{2}$', cell.replace(' ', '')):
                    continue
                # Skip dates
                if re.match(r'^\d{2}\s+\w{3}\s+\d{4}$', cell):
                    continue
                # Skip "Inactive Recipient"
                if 'Inactive' in cell:
                    continue

                # First valid text is likely the name
                if name is None and re.search(r'[A-Za-z]{2,}', cell):
                    name = clean_name(cell)
                elif name and reference is None and re.search(r'[A-Za-z]{2,}', cell):
                    reference = cell

            if name and account:
                recipients.append({
                    'name': name,
                    'account': account,
                    'reference': reference or name
                })

    return recipients


def extract_with_pymupdf(pdf_path: str) -> list[dict]:
    """Extract recipient data using PyMuPDF (fitz) with block extraction."""
    if fitz is None:
        raise ImportError("PyMuPDF is not installed")

    recipients = []

    doc = fitz.open(pdf_path)
    for page in doc:
        # Extract text blocks which preserve layout better
        blocks = page.get_text("dict")["blocks"]
        page_recipients = parse_pymupdf_blocks(blocks)
        recipients.extend(page_recipients)
    doc.close()

    return recipients


def parse_pymupdf_blocks(blocks: list) -> list[dict]:
    """Parse text blocks from PyMuPDF to extract recipient data."""
    recipients = []
    account_pattern = re.compile(r'^(\d{8,11})$')

    # Collect all text spans with their positions
    text_items = []
    for block in blocks:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    text_items.append({
                        "text": text,
                        "x": span["bbox"][0],
                        "y": span["bbox"][1],
                        "y_end": span["bbox"][3]
                    })

    # Sort by y position (top to bottom), then x position (left to right)
    text_items.sort(key=lambda item: (round(item["y"] / 10) * 10, item["x"]))

    # Group items by approximate y position (same row)
    rows = []
    current_row = []
    last_y = None

    for item in text_items:
        if last_y is None or abs(item["y"] - last_y) < 15:
            current_row.append(item)
        else:
            if current_row:
                rows.append(current_row)
            current_row = [item]
        last_y = item["y"]

    if current_row:
        rows.append(current_row)

    # Process rows to find recipients
    # Look for rows that contain an account number
    for row in rows:
        row_texts = [item["text"] for item in row]
        row_combined = " ".join(row_texts)

        # Skip headers and non-data rows
        skip_keywords = [
            'Name', 'Pay Amount', 'Last Paid', 'Their Reference', 'My Reference',
            'Please note', 'Due to system', 'Real-time', 'Education',
            'Entertainment', 'Medical', 'Motoring', 'Personal Services',
            'Household', 'Family', 'Not Categorised', 'View the cut-off',
            'Amount My'
        ]
        if any(keyword in row_combined for keyword in skip_keywords):
            continue

        # Find account numbers in row
        for i, text in enumerate(row_texts):
            if account_pattern.match(text):
                account = text
                # Name is typically in the text items before the account number
                name_parts = []
                for j in range(i):
                    t = row_texts[j]
                    # Skip monetary amounts and dates
                    if re.match(r'^[\d,]+\.\d{2}$', t.replace(' ', '')):
                        continue
                    if re.match(r'^\d{2}\s+\w{3}\s+\d{4}$', t):
                        continue
                    if t in ['0.00', 'Inactive Recipient']:
                        continue
                    name_parts.append(t)

                if name_parts:
                    name = ' '.join(name_parts)
                    name = clean_name(name)

                    # Reference is typically after the account number
                    ref_parts = []
                    for j in range(i + 1, len(row_texts)):
                        t = row_texts[j]
                        if re.match(r'^[\d,]+\.\d{2}$', t.replace(' ', '')):
                            continue
                        if re.match(r'^\d{2}\s+\w{3}\s+\d{4}$', t):
                            continue
                        ref_parts.append(t)

                    reference = ' '.join(ref_parts) if ref_parts else name

                    if name and account:
                        recipients.append({
                            'name': name,
                            'account': account,
                            'reference': reference
                        })
                break

    return recipients


def extract_with_ocr(pdf_path: str) -> list[dict]:
    """Extract recipient data using OCR (pytesseract)."""
    if pytesseract is None or Image is None:
        raise ImportError("pytesseract or Pillow is not installed")
    if fitz is None:
        raise ImportError("PyMuPDF is required for OCR extraction")

    recipients = []

    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc[page_num]
        # Render page to image at high resolution
        mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better OCR
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))

        # Perform OCR
        text = pytesseract.image_to_string(img)
        if text:
            recipients.extend(parse_fnb_text(text))

    doc.close()

    return recipients


def parse_fnb_text(text: str) -> list[dict]:
    """
    Parse FNB recipient text and extract recipient information.

    FNB format typically shows:
    - Beneficiary Name (e.g., "P Holroyd T/a Tiny Twiste")
    - Account Number (8-11 digit number)
    - Reference information
    """
    recipients = []

    # Pattern to match account numbers (8-11 digits typically for SA banks)
    account_pattern = re.compile(r'\b(\d{8,11})\b')

    # Split text into lines for processing
    lines = text.split('\n')

    # Track current recipient being built
    current_name = None
    current_account = None
    current_reference = None

    # Known section headers to skip
    skip_sections = [
        'Pay recipient', 'Due to system', 'Please note', 'Last Paid',
        'Pay Amount', 'Their Reference', 'My Reference', 'Name',
        'Education', 'Entertainment/sports', 'Medical', 'Motoring',
        'Personal Services', 'Household Maintenance', 'Family And Friends',
        'Not Categorised', 'Real-time payments', 'View the cut-off',
        'Inactive Recipient', 'Amount'
    ]

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Skip empty lines and section headers
        if not line or any(skip in line for skip in skip_sections):
            i += 1
            continue

        # Skip monetary amounts
        if re.match(r'^[\d,]+\.\d{2}$', line.replace(' ', '')):
            i += 1
            continue

        # Skip date patterns
        if re.match(r'^\d{2}\s+\w{3}\s+\d{4}$', line):
            i += 1
            continue

        # Look for account numbers
        account_match = account_pattern.search(line)
        if account_match:
            potential_account = account_match.group(1)
            # Check if this line is primarily an account number
            if line == potential_account or len(line) < 20:
                if current_name and not current_account:
                    current_account = potential_account
                elif current_name and current_account:
                    # Save previous recipient and start new one
                    recipients.append({
                        'name': current_name,
                        'account': current_account,
                        'reference': current_reference or current_name
                    })
                    current_name = None
                    current_account = potential_account
                    current_reference = None
                else:
                    current_account = potential_account
                i += 1
                continue

        # Check if this looks like a name (contains letters, possibly business suffixes)
        if re.search(r'[A-Za-z]{2,}', line) and not re.match(r'^\d+$', line):
            # This might be a name
            if current_name and current_account:
                # Save the previous recipient
                recipients.append({
                    'name': current_name,
                    'account': current_account,
                    'reference': current_reference or current_name
                })
                current_reference = None

            # Check if line contains both name and account
            if account_match:
                # Extract name part (before account number)
                name_part = line[:account_match.start()].strip()
                if name_part:
                    current_name = clean_name(name_part)
                    current_account = account_match.group(1)
                else:
                    current_name = clean_name(line.replace(account_match.group(1), '').strip())
                    current_account = account_match.group(1)
            else:
                current_name = clean_name(line)
                current_account = None

        i += 1

    # Don't forget the last recipient
    if current_name and current_account:
        recipients.append({
            'name': current_name,
            'account': current_account,
            'reference': current_reference or current_name
        })

    return recipients


def clean_name(name: str) -> str:
    """Clean up beneficiary name."""
    # Remove multiple spaces
    name = re.sub(r'\s+', ' ', name)
    # Remove leading/trailing whitespace
    name = name.strip()
    # Remove trailing numbers that might be partial account numbers
    name = re.sub(r'\s+\d+$', '', name)
    return name


def deduplicate_recipients(recipients: list[dict]) -> list[dict]:
    """Remove duplicate recipients based on account number."""
    seen_accounts = set()
    unique_recipients = []

    for recipient in recipients:
        account = recipient['account']
        if account not in seen_accounts:
            seen_accounts.add(account)
            unique_recipients.append(recipient)

    return unique_recipients


def convert_to_investec_format(recipients: list[dict]) -> list[dict]:
    """Convert FNB recipients to Investec CSV format."""
    investec_records = []

    for recipient in recipients:
        record = {
            "Beneficiary Account Name": recipient['name'],
            "Beneficiary Bank": "FNB",
            "Beneficiary Bank Account Number": recipient['account'],
            "Beneficiary Branch Code": FNB_BRANCH_CODE,
            "Beneficiary Reference": recipient.get('reference', recipient['name'])[:20],
            "Statement Description": recipient['name'][:20],
            "Beneficiary Name": recipient['name'],
            "Beneficiary Fax Number": "",
            "Beneficiary Email Address": "",
            "Beneficiary Cell Number": "",
        }
        investec_records.append(record)

    return investec_records


def write_investec_csv(records: list[dict], output_path: str) -> None:
    """Write records to Investec CSV format."""
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=INVESTEC_HEADERS)
        writer.writeheader()
        writer.writerows(records)


def extract_recipients(pdf_path: str, method: str = "auto") -> list[dict]:
    """
    Extract recipients from PDF using specified method.

    Args:
        pdf_path: Path to the FNB recipients PDF
        method: Extraction method - "pdfplumber", "pymupdf", "ocr", or "auto"

    Returns:
        List of recipient dictionaries
    """
    methods_to_try = []

    if method == "auto":
        # Try methods in order of preference
        if pdfplumber is not None:
            methods_to_try.append(("pdfplumber", extract_with_pdfplumber))
        if fitz is not None:
            methods_to_try.append(("pymupdf", extract_with_pymupdf))
        if pytesseract is not None and fitz is not None:
            methods_to_try.append(("ocr", extract_with_ocr))
    elif method == "pdfplumber":
        methods_to_try.append(("pdfplumber", extract_with_pdfplumber))
    elif method == "pymupdf":
        methods_to_try.append(("pymupdf", extract_with_pymupdf))
    elif method == "ocr":
        methods_to_try.append(("ocr", extract_with_ocr))
    else:
        raise ValueError(f"Unknown extraction method: {method}")

    if not methods_to_try:
        raise RuntimeError(
            "No PDF extraction libraries available. "
            "Please install pdfplumber, PyMuPDF, or pytesseract+Pillow."
        )

    recipients = []
    for method_name, extract_func in methods_to_try:
        try:
            print(f"Trying extraction method: {method_name}")
            recipients = extract_func(pdf_path)
            if recipients:
                print(f"Successfully extracted {len(recipients)} recipients using {method_name}")
                break
            else:
                print(f"No recipients found using {method_name}, trying next method...")
        except Exception as e:
            print(f"Method {method_name} failed: {e}")
            continue

    return recipients


def main():
    parser = argparse.ArgumentParser(
        description="Convert FNB recipients PDF to Investec CSV format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s fnb-recipients.pdf
  %(prog)s fnb-recipients.pdf -o my-beneficiaries.csv
  %(prog)s fnb-recipients.pdf --method ocr

Supported extraction methods:
  auto       - Automatically try available methods (default)
  pdfplumber - Use pdfplumber library
  pymupdf    - Use PyMuPDF (fitz) library
  ocr        - Use OCR via pytesseract (requires Tesseract installed)
        """
    )

    parser.add_argument(
        "pdf_file",
        help="Path to FNB recipients PDF file"
    )

    parser.add_argument(
        "-o", "--output",
        default="investec-beneficiaries.csv",
        help="Output CSV file path (default: investec-beneficiaries.csv)"
    )

    parser.add_argument(
        "-m", "--method",
        choices=["auto", "pdfplumber", "pymupdf", "ocr"],
        default="auto",
        help="PDF extraction method (default: auto)"
    )

    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output"
    )

    args = parser.parse_args()

    # Validate input file
    pdf_path = Path(args.pdf_file)
    if not pdf_path.exists():
        print(f"Error: PDF file not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    if not pdf_path.suffix.lower() == '.pdf':
        print(f"Warning: File does not have .pdf extension: {pdf_path}", file=sys.stderr)

    # Extract recipients
    print(f"Processing: {pdf_path}")
    recipients = extract_recipients(str(pdf_path), method=args.method)

    if not recipients:
        print("Error: No recipients could be extracted from the PDF.", file=sys.stderr)
        print("Try using a different extraction method with --method", file=sys.stderr)
        sys.exit(1)

    # Deduplicate
    original_count = len(recipients)
    recipients = deduplicate_recipients(recipients)
    if len(recipients) < original_count:
        print(f"Removed {original_count - len(recipients)} duplicate entries")

    # Convert to Investec format
    investec_records = convert_to_investec_format(recipients)

    # Write output
    write_investec_csv(investec_records, args.output)
    print(f"Successfully wrote {len(investec_records)} beneficiaries to: {args.output}")

    if args.verbose:
        print("\nExtracted beneficiaries:")
        for record in investec_records:
            print(f"  - {record['Beneficiary Account Name']} ({record['Beneficiary Bank Account Number']})")


if __name__ == "__main__":
    main()
