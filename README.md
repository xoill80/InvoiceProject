# Invoice Verification System

A Python-based invoice verification system that compares Purchase Order, Receipt, and Invoice data.

## Features
- PDF data extraction
- OCR image reading
- Quantity verification
- Price verification
- PO Number verification
- Supplier verification
- VAT verification
- Total verification
- Arabic and English discrepancy explanation
- Tkinter dashboard

## Technologies
- Python
- Tesseract OCR
- pytesseract
- pdfplumber
- pypdf
- openpyxl
- Tkinter

## Results
The system displays:

- MATCHED when all document values are consistent.
- REVIEW NEEDED when discrepancies are detected.

## Example Checks
- Quantity mismatch
- Price mismatch
- PO Number mismatch
- Supplier mismatch
- VAT mismatch
- Total mismatch

## How to Run
Run:

```bash
python app.py
'''