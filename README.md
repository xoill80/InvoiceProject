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


## Web Application

Invoice Verification is a web application that checks Purchase Orders (PO), Receipts, and Invoices for inconsistencies.

## Website

https://invoice-verification.streamlit.app

## How to Use

1. Upload your Purchase Order (PO).
2. Upload your Receipt.
3. Upload your Invoice.
4. Click Verify Documents.
5. Review the verification result.
6. Click Download Report to save the results.

## Verification Results

1. MATCHED: The checked document values are consistent.
2. REVIEW NEEDED: One or more differences were detected and should be reviewed.

## Important Notes

1. Supported file types include PDF, XLSX, PNG, and JPG.
2. This application is currently a prototype undergoing testing.
3. Use sample or non-sensitive documents during testing.
4. Always review the results before making financial decisions.

