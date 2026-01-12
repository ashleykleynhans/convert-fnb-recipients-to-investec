# FNB Recipients to Investec CSV Converter

A Python tool to convert FNB (First National Bank) recipient PDF exports to Investec's
beneficiary import CSV format.

## Overview

When switching from FNB to Investec, you may want to import your existing payment recipients.
FNB allows you to export your recipients as a PDF, but Investec requires a specific CSV format
for importing beneficiaries. This tool bridges that gap by parsing the FNB PDF and generating
an Investec-compatible CSV file.

## Features

- Parses FNB recipient PDF exports using multiple extraction methods
- Supports pdfplumber, PyMuPDF, and OCR (pytesseract) extraction
- Automatically deduplicates recipients
- Generates CSV in Investec's beneficiary import format
- Uses FNB's universal branch code (250655) for all recipients
- Command-line interface with verbose output option

## Requirements

- Python 3.10+
- For OCR support: Tesseract OCR installed on your system

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/convert-fnb-recipients-to-investec.git
   cd convert-fnb-recipients-to-investec
   ```

2. Create and activate a virtual environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. (Optional) For OCR support, install Tesseract:
   - **macOS**: `brew install tesseract`
   - **Ubuntu/Debian**: `sudo apt-get install tesseract-ocr`
   - **Windows**: Download from [GitHub](https://github.com/UB-Mannheim/tesseract/wiki)

## Usage

### Basic Usage

```bash
python convert_fnb_to_investec.py fnb-recipients.pdf
```

This will create `investec-beneficiaries.csv` in the current directory.

### Options

```bash
python convert_fnb_to_investec.py fnb-recipients.pdf [OPTIONS]

Options:
  -o, --output FILE    Output CSV file path (default: investec-beneficiaries.csv)
  -m, --method METHOD  PDF extraction method: auto, pdfplumber, pymupdf, ocr (default: auto)
  -v, --verbose        Enable verbose output
  -h, --help           Show help message
```

### Examples

```bash
# Specify output file
python convert_fnb_to_investec.py fnb-recipients.pdf -o my-beneficiaries.csv

# Use specific extraction method
python convert_fnb_to_investec.py fnb-recipients.pdf --method pymupdf

# Verbose output showing all extracted beneficiaries
python convert_fnb_to_investec.py fnb-recipients.pdf -v
```

## Output Format

The generated CSV follows Investec's beneficiary import template:

| Column | Description |
|--------|-------------|
| Beneficiary Account Name | Account holder name |
| Beneficiary Bank | Bank name (FNB) |
| Beneficiary Bank Account Number | Account number |
| Beneficiary Branch Code | Universal branch code (250655) |
| Beneficiary Reference | Payment reference (max 20 chars) |
| Statement Description | Statement description (max 20 chars) |
| Beneficiary Name | Full beneficiary name |
| Beneficiary Fax Number | (empty) |
| Beneficiary Email Address | (empty) |
| Beneficiary Cell Number | (empty) |

## How to Export Recipients from FNB

1. Log in to FNB Online Banking
2. Navigate to **Payments** > **Pay recipient**
3. Use your browser's print function (Ctrl+P / Cmd+P)
4. Select "Save as PDF"
5. Save the PDF file

## Extraction Methods

The tool supports multiple PDF extraction methods:

- **pdfplumber** (default): Best for most FNB PDFs, uses word-level extraction with position data
- **pymupdf**: Alternative text block extraction using PyMuPDF
- **ocr**: Uses Tesseract OCR for scanned or image-based PDFs

The `auto` method tries each in order until one succeeds.

## Limitations

- Only extracts recipients with valid FNB account numbers (8-11 digits)
- Email, fax, and cell number fields are left empty (not available in FNB PDF)
- All recipients are assumed to be FNB accounts with branch code 250655
- Reference fields are truncated to 20 characters (Investec limit)

## Troubleshooting

### No recipients found

Try a different extraction method:
```bash
python convert_fnb_to_investec.py fnb-recipients.pdf --method ocr
```

### Garbled or incorrect names

The PDF may have a complex layout. Try the PyMuPDF method:
```bash
python convert_fnb_to_investec.py fnb-recipients.pdf --method pymupdf
```

### OCR not working

Ensure Tesseract is installed and in your PATH:
```bash
tesseract --version
```

## License

This project is licensed under the GNU General Public License v3.0
- see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Disclaimer

This tool is provided as-is for personal use. Always verify the generated CSV
before importing into Investec. The authors are not responsible for any errors
in the conversion process or any issues arising from the use of this tool.
