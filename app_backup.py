import json
import csv
import re
from openpyxl import load_workbook
from pypdf import PdfReader
import pdfplumber

def read_receipt_pdf():
    receipt_pdf_reader = PdfReader("receipt.pdf")
    receipt_pdf_text = ""

    for page in receipt_pdf_reader.pages:
        receipt_pdf_text = receipt_pdf_text + page.extract_text()

    with pdfplumber.open("receipt.pdf") as pdf:
        receipt_first_page = pdf.pages[0]
        receipt_words = receipt_first_page.extract_words()

    receipt_word_texts = [word["text"] for word in receipt_words]
    pdf_receipt_quantity = int(receipt_word_texts[1])

    return {
        "quantity": pdf_receipt_quantity
    }        

def read_po_pdf():
    with pdfplumber.open("po.pdf") as pdf:
        po_first_page = pdf.pages[0] 
        po_words = po_first_page.extract_words()

    po_word_texts = [word["text"] for word in po_words]

    pdf_po_number_from_po = po_word_texts[4]
    pdf_po_supplier = po_word_texts[5] + " " + po_word_texts[6]
    pdf_po_quantity = int(po_word_texts[7])
    pdf_po_price = float(po_word_texts[8])

    return {
        "number": pdf_po_number_from_po,
        "supplier": pdf_po_supplier,
        "quantity": pdf_po_quantity,
        "price": pdf_po_price
    }    

def read_invoice_pdf():
    pdf_reader = PdfReader("invoice.pdf")
    pdf_text = ""

    for page in pdf_reader.pages:
        pdf_text = pdf_text + page.extract_text()

    po_match = re.search(r"PO-\d+", pdf_text)

    if po_match:
        pdf_po_number = po_match.group()
    else:
        pdf_po_number = ""

    supplier_match = re.search(r"([A-Za-z][A-Za-z ]+?)\s*PO-\d+", pdf_text)

    if supplier_match:
        pdf_supplier = supplier_match.group(1).strip()
    else:
        pdf_supplier = ""

    with pdfplumber.open("invoice.pdf") as pdf:
        first_page = pdf.pages[0]
        words = first_page.extract_words()

    word_texts = [word["text"] for word in words]

    pdf_quantity = int(word_texts[10])
    pdf_price = float(word_texts[11])
    pdf_subtotal = float(word_texts[12])
    pdf_vat = float(word_texts[13])
    pdf_total = float(word_texts[14])

    return {
        "po_number": pdf_po_number,
        "supplier": pdf_supplier,
        "quantity": pdf_quantity,
        "price": pdf_price,
        "subtotal": pdf_subtotal,
        "vat": pdf_vat,
        "total": pdf_total
    }                        

invoice_from_pdf = read_invoice_pdf()
invoice = invoice_from_pdf

with open("sample_data.json", "r", encoding="utf-8") as file:
    data = json.load(file)

invoice_from_pdf = read_invoice_pdf()

receipt_from_pdf = read_receipt_pdf()
po_from_pdf = read_po_pdf()

receipt = receipt_from_pdf
po = po_from_pdf
invoice = invoice_from_pdf

def check_quantity(po_quantity, received_quantity, invoice_quantity):
    if po_quantity == received_quantity == invoice_quantity:
        print("Quantity: MATCH")
        return 0

    print("Quantity: MISMATCH")
    print("Ordered:", po_quantity)
    print("Received:", received_quantity)
    print("Invoiced", invoice_quantity)
    print("Difference", invoice_quantity - received_quantity)
    return 1


def check_price(po_price, invoice_price):
    if po_price == invoice_price:
        print("Price: MATCH")
        return 0

    print("Price: MISMATCH")
    print("PO Price:", po_price)
    print("Invoice Price:", invoice_price)
    print("Price Difference", invoice_price - po_price)
    return 1


def check_po_number(po_number, invoice_po_number):
    clean_po_number = po_number.strip().upper()
    clean_invoice_po_number = invoice_po_number.strip().upper()

    if clean_po_number == clean_invoice_po_number:
        print("PO Number: MATCH")
        return 0

    print("PO Number: MISMATCH")
    print("PO Number:", po_number)
    print("Invoice PO Number:", invoice_po_number)
    return 1


def check_supplier(po_supplier, invoice_supplier):
    clean_po_supplier = po_supplier.strip().upper()
    clean_invoice_supplier = invoice_supplier.strip().upper()

    if clean_po_supplier == clean_invoice_supplier:
        print("Supplier: MATCH")
        return 0

    print("Supplier: MISMATCH") 
    print("PO Supplier:", po_supplier)  
    print("Invoice Supplier:", invoice_supplier) 
    return 1


def check_vat(subtotal, invoice_vat):
    expected_vat = subtotal * 0.15

    if expected_vat == invoice_vat:
        print("VAT: MATCH")
        return 0

    print("VAT: MISMATCH")
    print("Expected VAT:", expected_vat)
    print("Invoice VAT:", invoice_vat)
    return 1


def check_total(subtotal, invoice_vat, invoice_total):
    expected_total = subtotal + invoice_vat

    if expected_total == invoice_total:
        print("Total: MATCH")
        return 0

    print("Total: MISMATCH")
    print("Expected Total:", expected_total)
    print("Invoice Total:", invoice_total)
    return 1


issues_found = 0

issues_found = issues_found + check_quantity(
    po["quantity"],
    receipt["quantity"],
    invoice["quantity"]
)

issues_found = issues_found + check_price(
    po["price"],
    invoice["price"]
)

issues_found = issues_found + check_po_number(
    po["number"],
    invoice["po_number"]
)

issues_found = issues_found + check_supplier(
    po["supplier"],
    invoice["supplier"]
)

issues_found = issues_found + check_vat(
   invoice["subtotal"],
    invoice["vat"]
)

issues_found = issues_found + check_total(
    invoice["subtotal"],
    invoice["vat"],
    invoice["total"]
)

if issues_found == 0:
    print("FINAL RESULT: MATCHED") 
else:
    print("FINAL RESULT: REVIEW NEEDED") 
    print("Issues Found:", issues_found)            



