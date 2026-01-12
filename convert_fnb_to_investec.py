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

    # For each account number, find the associated name and reference
    # Names are ABOVE the account number in FNB PDFs (within ~35 pixels above)
    # References are in the rightmost column (x > 400)
    for acc_info in account_positions:
        account = acc_info['account']
        acc_y = acc_info['y']

        name_words = []
        ref_words = []

        for word in words_sorted:
            word_y = word['top']
            word_x = word['x0']
            text = word['text'].strip()

            if not text or text == account:
                continue

            # Check if word is above the account (within 35 pixels)
            y_diff = acc_y - word_y
            if 0 < y_diff <= 35:
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

                # Name column (x < 130)
                if word_x < 130:
                    name_words.append({'text': text, 'x': word_x, 'y': word_y})
                # Reference column (x > 400, excluding amounts)
                elif word_x > 400:
                    ref_words.append({'text': text, 'x': word_x, 'y': word_y})

        # Build name
        if name_words:
            name_words.sort(key=lambda w: (w['y'], w['x']))
            name = ' '.join(w['text'] for w in name_words)
            name = clean_name(name)

            # Build reference - deduplicate words at same position
            reference = name  # Default to name
            if ref_words:
                # Sort by y then x, and deduplicate overlapping text
                ref_words.sort(key=lambda w: (w['y'], w['x']))
                seen_positions = set()
                unique_ref_words = []
                for w in ref_words:
                    pos_key = (round(w['y'] / 5), round(w['x'] / 5))
                    if pos_key not in seen_positions:
                        seen_positions.add(pos_key)
                        unique_ref_words.append(w['text'])
                if unique_ref_words:
                    reference = ' '.join(unique_ref_words)

            if name and len(name) > 1:
                recipients.append({
                    'name': name,
                    'account': account,
                    'reference': reference
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
    """Extract recipient data using PyMuPDF (fitz) with block-level extraction."""
    if fitz is None:
        raise ImportError("PyMuPDF is not installed")

    recipients = []

    doc = fitz.open(pdf_path)
    for page in doc:
        page_data = page.get_text("dict")
        page_recipients = parse_pymupdf_blocks_v2(page_data["blocks"])
        recipients.extend(page_recipients)
    doc.close()

    return recipients


def parse_pymupdf_blocks_v2(blocks: list) -> list[dict]:
    """
    Parse text blocks from PyMuPDF to extract recipient data.
    FNB PDF has overlapping columns - 'Their Reference' and 'My Reference' at same x position.
    We extract the second occurrence which is 'My Reference'.
    """
    recipients = []
    account_pattern = re.compile(r'^(\d{8,11})$')

    # Collect all spans with block index to track overlapping blocks
    all_spans = []
    for block_idx, block in enumerate(blocks):
        if "lines" not in block:
            continue
        block_x = block["bbox"][0]
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if text:
                    all_spans.append({
                        "text": text,
                        "x": span["bbox"][0],
                        "y": span["bbox"][1],
                        "block_idx": block_idx,
                        "block_x": block_x
                    })

    # Find account numbers
    account_spans = [s for s in all_spans if account_pattern.match(s["text"])]

    for acc_span in account_spans:
        account = acc_span["text"]
        acc_y = acc_span["y"]

        # Find name - spans above account (within 35px), in left column (x < 130)
        name_spans = [
            s for s in all_spans
            if s["x"] < 130
            and 0 < acc_y - s["y"] <= 35
            and not re.match(r'^[\d,]+\.\d{2}$', s["text"].replace(' ', ''))
            and s["text"] not in ['0.00', 'Inactive', 'Recipient']
        ]
        name_spans.sort(key=lambda s: (s["y"], s["x"]))
        name = ' '.join(s["text"] for s in name_spans)
        name = clean_name(name)

        # Find reference - spans in reference column (x > 430), above account
        # Use y-range of 40px to capture multi-line references without bleeding into adjacent rows
        # Filter out header text
        ref_spans = [
            s for s in all_spans
            if s["x"] > 430
            and 0 < acc_y - s["y"] <= 40
            and not re.match(r'^[\d,]+\.\d{2}$', s["text"].replace(' ', ''))
            and s["text"].strip() not in ['Their', 'My', 'Reference', 'Amount']
        ]

        # Group spans by y-position (rounded to nearest 2px for more precise grouping)
        # At each y-level, there may be two spans: 'Their Reference' then 'My Reference'
        # We want the second one (My Reference)
        ref_by_y = {}
        for s in ref_spans:
            y_key = round(s["y"] / 2) * 2  # Round to nearest 2px
            if y_key not in ref_by_y:
                ref_by_y[y_key] = []
            ref_by_y[y_key].append(s)

        # Build reference from 'My Reference' column
        # FNB PDFs have overlapping text - second span at each y is 'My Reference'
        reference = name
        if ref_by_y:
            my_ref_spans = []
            for y_key in sorted(ref_by_y.keys()):
                spans_at_y = ref_by_y[y_key]
                if len(spans_at_y) >= 2:
                    # Two columns - take the second (My Reference)
                    my_ref_spans.append(spans_at_y[1])
                elif len(spans_at_y) == 1:
                    # Single span - check if it looks like a continuation
                    # or if it might be part of a split reference
                    span = spans_at_y[0]
                    my_ref_spans.append(span)
            if my_ref_spans:
                # Join and clean up the reference
                ref_text = ' '.join(s["text"].strip() for s in my_ref_spans)
                ref_text = re.sub(r'\s+', ' ', ref_text).strip()
                if ref_text:
                    reference = ref_text

        if name and len(name) > 1:
            recipients.append({
                'name': name,
                'account': account,
                'reference': reference
            })

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
        # PyMuPDF first as it handles FNB PDF column structure better
        if fitz is not None:
            methods_to_try.append(("pymupdf", extract_with_pymupdf))
        if pdfplumber is not None:
            methods_to_try.append(("pdfplumber", extract_with_pdfplumber))
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
