import json
import csv
import re

from openpyxl import load_workbook
from pypdf import PdfReader
import pdfplumber
import pytesseract
from PIL import Image

try:
   import tkinter as tk
   from tkinter import filedialog, messagebox
except ImportError:
    tk = None
    filedialog = None
    messagebox = None   

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def test_ocr():
    image = Image.open("test_invoice.png")
    text = pytesseract.image_to_string(image)

    print("===== OCR TEXT =====")
    print(text)

def read_invoice_image(file_path):
    image = Image.open(file_path)

    text = pytesseract.image_to_string(image)

    po_match = re.search(r"P[O0]\s*-\s*\d+", text, re.IGNORECASE)

    if po_match:
        po_number = po_match.group().upper()
        po_number = po_number.replace("P0", "PO").replace(" ", "")
    else:
        po_number = "" 

    supplier_match = re.search(r"Supplier:\s*(.+)", text)

    if supplier_match:
        supplier = supplier_match.group(1).strip()
    else:
        supplier = ""

    quantity_match = re.search(r"Quantity:\s*(\d+)", text)

    if quantity_match:
        quantity = int(quantity_match.group(1))
    else:
        quantity = None

    price_match = re.search(r"Price:\s*([0-90o]+(?:\.[0-9]+)?)", text)

    if price_match:
        price_text = price_match.group(1).replace("O", "0").replace("O", "0")
        price = float(price_text) 
    else:
        price = None

    subtotal_match = re.search(r"Subtotal:\s*([0-90o]+(?:\.[0-9]+)?)", text)

    if subtotal_match:
        subtotal_text = subtotal_match.group(1).replace("O", "0").replace("o", "0")
        subtotal = float(subtotal_text)
    else:
        subtotal = None

    if quantity is not None and price is not None:
        calculated_subtotal = quantity * price

        if subtotal is None or subtotal != calculated_subtotal:
            subtotal = calculated_subtotal

    vat_match = re.search(r"VAT:\s*([0-90o]+(?:\.[0-9]+)?)", text)

    if vat_match:
        vat_text = vat_match.group(1).replace("O", "0").replace("o", "0")
        vat = float(vat_text)
    else:
        vat = None

    total_match = re.search(r"Total:\s*([0-90o]+(?:\.[0-9]+)?)", text)

    if subtotal is not None and vat is not None:
        total = subtotal + vat
    else:
        total = None                                           

    return {
        "po_number": po_number,
        "supplier": supplier,
        "quantity": quantity,
        "price": price,
        "subtotal": subtotal,
        "vat": vat,
        "total": total
    }          


# invoice_image = read_invoice_image()


def read_receipt_pdf(file_path="receipt.pdf"):
    receipt_pdf_reader = PdfReader(file_path)
    receipt_pdf_text = ""

    for page in receipt_pdf_reader.pages:
        receipt_pdf_text = receipt_pdf_text + page.extract_text()

    with pdfplumber.open(file_path) as pdf:
        receipt_first_page = pdf.pages[0]
        receipt_words = receipt_first_page.extract_words()

    receipt_word_texts = [word["text"] for word in receipt_words]

    pdf_receipt_quantity = None

    for i, text in enumerate(receipt_word_texts):
        if text.lower() == "quantity" and i + 1 < len(receipt_word_texts):
            pdf_receipt_quantity = int(receipt_word_texts[i + 1])
            break

    return {
        "quantity": pdf_receipt_quantity
    }  


def read_receipt_image(file_path):
    image = Image.open(file_path)

    image = image.convert("L")
    image = image.resize((image.width * 3, image.height * 3))

    text = pytesseract.image_to_string(
        image,
        config="--psm 6"
    )


    quantity_match = re.search(
        r"quantity\s*[:\-]?\s*(\d+)", 
        text, 
        re.IGNORECASE
    )

    if quantity_match:
        quantity = int(quantity_match.group(1))
    else:
        quantity = None

    return {
        "quantity": quantity
    }        


def read_receipt_excel(file_path):
    workbook = load_workbook(file_path, data_only=True)
    sheet = workbook.active

    quantity = sheet["A2"].value

    return {
        "quantity": quantity
    }


def read_po_pdf(file_path="po.pdf"):
    with pdfplumber.open(file_path) as pdf:
        po_first_page = pdf.pages[0] 
        po_words = po_first_page.extract_words()

    po_word_texts = [word["text"] for word in po_words]

    pdf_po_number_from_po = po_word_texts[4]
    pdf_po_supplier = po_word_texts[5] + " " + po_word_texts[6]

    pdf_po_quantity = None

    for i, text in enumerate(po_word_texts):
        if text.lower() == "quantity":
            for value in po_word_texts[i + 1:]:
                if value.isdigit():
                    pdf_po_quantity = int(value)
                    break
            break 

    pdf_po_price = None
    numbers_after_price = []

    for i, text in enumerate(po_word_texts):
        if text.lower() == "price":
            for value in po_word_texts[i + 1:]:
                try:
                    numbers_after_price.append(float(value))
                except ValueError:
                    continue
            break

    if len(numbers_after_price) >= 2:
        pdf_po_price = numbers_after_price[1]            
       
                   
    return {
        "number": pdf_po_number_from_po,
        "supplier": pdf_po_supplier,
        "quantity": pdf_po_quantity,
        "price": pdf_po_price
    } 


def read_po_image(file_path):
    image = Image.open(file_path)

    image = image.convert("L")
    image = image.resize((image.width * 3, image.height * 3))

    text = pytesseract.image_to_string(
        image,
        config="--psm 6"
    )


    po_match = re.search(
        r"PO[-\s]?\d+",
        text,
        re.IGNORECASE
    )

    supplier_match = re.search(
        r"Supplier\s*[:\-]?\s*(.+)",
        text,
        re.IGNORECASE
    )

    quantity_match = re.search(
        r"Quantity\s*[:\-]?\s*([0-9Oo]+)",
        text,
        re.IGNORECASE
    )

    price_match = re.search(
        r"Price\s*[:\-]?\s*([0-9Oo]+(?:\.[0-9Oo]+)?)",
        text,
        re.IGNORECASE
    )

    if po_match:
        po_number = po_match.group().replace(" ", "").upper()
    else:
        po_number = ""

    if supplier_match:
        supplier = supplier_match.group(1).strip()
    else:
        supplier = ""

    if quantity_match:
        quantity_text = quantity_match.group(1).replace("O", "0").replace("o", "0")
        quantity = int(quantity_text)
    else:
        quantity = None

    if price_match:
        price_text = price_match.group(1).replace("O", "0").replace("o", "0")
        price = float(price_text)
    else:
        price = None

    return {
        "number": po_number,
        "supplier": supplier,
        "quantity": quantity,
        "price": price
    }

def read_po_excel(file_path):
    workbook = load_workbook(file_path, data_only=True)
    sheet = workbook.active

    po_number = sheet["D2"].value
    supplier = sheet["C2"].value
    quantity = sheet["B2"].value
    price = sheet["A2"].value

    return {
        "number": po_number,
        "supplier": supplier,
        "quantity": quantity,
        "price": price
    }


def read_invoice_pdf(file_path="invoice.pdf"):
    pdf_reader = PdfReader(file_path)
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

    with pdfplumber.open(file_path) as pdf:
        first_page = pdf.pages[0]
        words = first_page.extract_words()

    word_texts = [word["text"] for word in words]

    pdf_quantity = None

    for i, text in enumerate(word_texts):
        if text.lower() == "quantity":
            for value in word_texts[i + 1:]:
                if value.isdigit():
                    pdf_quantity = int(value)
                    break
            break

    pdf_price = None
    numbers_after_price = []

    for i, text in enumerate(word_texts):
        if text.lower() == "price":
            for value in word_texts[i + 1:]:
                try:
                    numbers_after_price.append(float(value))
                except ValueError:
                    continue
            break

    if len(numbers_after_price) >= 2:
        pdf_price = numbers_after_price[1]                

    pdf_subtotal = None
    numbers_after_subtotal = []
    
    for i, text in enumerate(word_texts):
        if text.lower() == "subtotal":
            for value in word_texts[i + 1:]:
                try:
                    numbers_after_subtotal.append(float(value))
                except ValueError:
                    continue
            break

    if len(numbers_after_subtotal) >= 3:
        pdf_subtotal = numbers_after_subtotal[2]    
   
    pdf_vat = None
    numbers_after_vat = []
    
    for i, text in enumerate(word_texts):
        if text.lower() == "vat":
            for value in word_texts[i + 1:]:
                try:
                    numbers_after_vat.append(float(value))
                except ValueError:
                    continue
            break
    
    if len(numbers_after_vat) >= 4:
        pdf_vat = numbers_after_vat[3]

    pdf_total = None
    numbers_after_total = []
    
    for i, text in enumerate(word_texts):
        if text.lower() == "total":
            for value in word_texts[i + 1:]:
                try:
                    numbers_after_total.append(float(value))
                except ValueError:
                    continue
            break
    
    if len(numbers_after_total) >= 5:
        pdf_total = numbers_after_total[4]

    return {
        "po_number": pdf_po_number,
        "supplier": pdf_supplier,
        "quantity": pdf_quantity,
        "price": pdf_price,
        "subtotal": pdf_subtotal,
        "vat": pdf_vat,
        "total": pdf_total
    }                        

with open("sample_data.json", "r", encoding="utf-8") as file:
    data = json.load(file)


def read_invoice_excel(file_path):
    workbook = load_workbook(file_path, data_only=True) 
    sheet = workbook.active

    po_number = sheet["G2"].value
    supplier = sheet["F2"].value
    quantity = sheet["E2"].value
    price = sheet["D2"].value
    subtotal = sheet["C2"].value
    vat = sheet["B2"].value
    total = sheet["A2"].value

    return {
        "po_number": po_number,
        "supplier": supplier,
        "quantity": quantity,
        "price": price,
        "subtotal": subtotal,
        "vat": vat,
        "total": total
    }   


def normalize_invoice(invoice_data):
    return {
        "po_number": invoice_data.get("po_number", ""),
        "supplier": invoice_data.get("supplier", ""),
        "quantity": invoice_data.get("quantity"),
        "price": invoice_data.get("price"),
        "subtotal": invoice_data.get("subtotal"),
        "vat": invoice_data.get("vat"),
        "total": invoice_data.get("total")
    }                


def normalize_po(po_data):
    return {
        "number": po_data.get("number", ""),
        "supplier": po_data.get("supplier", ""),
        "quantity": po_data.get("quantity"),
        "price": po_data.get("price")
    }


def normalize_receipt(receipt_data):
    return {
        "quantity": receipt_data.get("quantity")
    }


def normalize_all(po_data, receipt_data, invoice_data):
    po = normalize_po(po_data)
    receipt = normalize_receipt(receipt_data)
    invoice = normalize_invoice(invoice_data)

    return po, receipt, invoice  


def check_quantity(po_quantity, received_quantity, invoice_quantity):
    if po_quantity is None or received_quantity is None or invoice_quantity is None:
        return 1
    
    if po_quantity == received_quantity == invoice_quantity:
        return 0

    return 1


def check_price(po_price, invoice_price):
    if po_price == invoice_price:
        return 0

    return 1


def check_po_number(po_number, invoice_po_number):
    clean_po_number = po_number.strip().upper()
    clean_invoice_po_number = invoice_po_number.strip().upper()

    if clean_po_number == clean_invoice_po_number:
        return 0

    return 1


def check_supplier(po_supplier, invoice_supplier):
    clean_po_supplier = po_supplier.strip().upper()
    clean_invoice_supplier = invoice_supplier.strip().upper()

    if clean_po_supplier == clean_invoice_supplier:
        return 0
  
    return 1


def check_vat(subtotal, invoice_vat):
    if subtotal is None or invoice_vat is None:
        return 1
    
    expected_vat = subtotal * 0.15

    if expected_vat == invoice_vat:
        return 0

    return 1


def check_total(subtotal, invoice_vat, invoice_total):
    if subtotal is None or invoice_vat is None or invoice_total is None:
        return 1
    
    expected_total = subtotal + invoice_vat

    if expected_total == invoice_total:
        return 0

    return 1


def print_summary(po, receipt, invoice, issues_found):

    problems = []

    if po["number"].strip().upper() != invoice["po_number"].strip().upper():
        problems.append(
            f"PO Number mismatch: po={po['number']} / Invoice={invoice['po_number']}"
        )

    if po["supplier"].strip().upper() != invoice["supplier"].strip().upper():
        problems.append(
            f"Supplier mismatch: PO={po['supplier']} / Invoice={invoice['supplier']}"
        )    

    if not (
        po["quantity"] == receipt["quantity"] == invoice["quantity"]
    ):
        problems.append(
            f"Quantity mismatch: Ordered={po['quantity']} / "
            f"Received={receipt['quantity']} / "
            f"Invoiced={invoice['quantity']}"
        ) 

    if po["price"] != invoice["price"]:
        problems.append(
            f"Price mismatch: PO={po['price']} / Invoice={invoice['price']}"
        )  

    if invoice["subtotal"] is not None and invoice["vat"] is not None:
        expected_vat = round(invoice["subtotal"] * 0.15, 2)

        if round(invoice["vat"], 2) != expected_vat:
            problems.append(
                 f"VAT mismatch: Excepted={expected_vat} / Invoice={invoice['vat']}"
            )
    else:
        problems.append("VAT data missing")


    if (
        invoice["subtotal"] is not None
        and invoice["vat"] is not None
        and invoice["total"]is not None
    ):    
        expected_total = round( invoice["subtotal"] + invoice["vat"], 2 ) 

        if round(invoice["total"], 2) != expected_total:
            problems.append(
                f"Total mismatch: Expected={expected_total} / Invoice={invoice['total']}"
            )   
    else:
        problems.append("Total data missing")  

    if issues_found == 0:
        pass
    else:
        pass


def explain_result(po, receipt, invoice, issues_found):
    if issues_found == 0:
        arabic_text = (
 "جميع بيانات أمر الشراء والاستلام والفاتورة متطابقة، "   
 ".ولم يتم العثور على أي اختلافات"
        ) 

        english_text = (
            "The PO, receipt, and invoice are consistent."
            "No discrepancies were found."
        )

        return arabic_text, english_text
    
    problems_ar = []
    problems_en = []

    if po["quantity"] != receipt["quantity"]or po["quantity"] != invoice["quantity"]:
        problems_ar.append("يوجد اختلاف في الكمية")
        problems_en.append("Quantity mismatch")

    if po["price"] != invoice["price"]:
        problems_ar.append("يوجد اختلاف في السعر")
        problems_en.append("Price mismatch")

    if po["number"] != invoice["po_number"]:
        problems_ar.append("يوجد اختلاف في رقم أمر الشراء")
        problems_en.append("PO number mismatch")

    if po["supplier"] != invoice["supplier"]:
        problems_ar.append("يوجد اختلاف في اسم المورد")
        problems_en.append("Supplier mismatch")

    if invoice["subtotal"] is None or invoice["vat"]is None:
        problems_ar.append("بيانات ضريبة القيمة المضافة غير مكتملة")
        problems_en.append("VAT data missing")
    else:
        expected_vat = round(invoice["subtotal"] * 0.15, 2)

        if round(invoice["vat"], 2) != expected_vat:
            problems_ar.append("يوجد اختلاف في ضريبة القيمة المضافة")
            problems_en.append("VAT mismatch")

    if (
        invoice["subtotal"] is None
        or invoice["vat"] is None
        or invoice["total"] is None
    ):
        problems_ar.append("بيانات الإجمالي غير مكتملة")
        problems_en.append("Total data missing")
    else:
        expected_total = round(
            invoice["subtotal"] + invoice["vat"],
            2
        )

        if round(invoice["total"], 2) != expected_total:
            problems_ar.append("يوجد اختلاف في المبلغ الإجمالي")
            problems_en.append("Total mismatch")

    arabic_text = "المشكلات المكتشفة: " + "، ".join(problems_ar)
    english_text = "Detected issue(s): " + ", ".join(problems_en)

    return arabic_text, english_text                                           


po_label = None

selected_po_file = ""

def select_po_file():
    global selected_po_file

    file_path = filedialog.askopenfilename(
        title="Select Purchase Order",
        filetypes=[
            ("PDF files", "*.pdf"),
            ("Image files", "*.png *.jpg *.jpeg"),
            ("Excel files", "*.xlsx")
        ]
    )

    if file_path:
        selected_po_file = file_path
        po_label.config(text=file_path.split("/")[-1])


receipt_label = None

selected_receipt_file = ""

def select_receipt_file():
    global selected_receipt_file

    file_path = filedialog.askopenfilename(
        title="Select Receipt",
        filetypes=[
            ("PDF files", "*.pdf"),
            ("Image files", "*.png *.jpg *.jpeg"),
            ("Excel files", "*.xlsx")
        ]
    )
    if file_path:
        selected_receipt_file = file_path
        receipt_label.config(text=file_path.split("/")[-1])


invoice_label = None

selected_invoice_file = "" 

def select_invoice_file():
    global selected_invoice_file

    file_path = filedialog.askopenfilename(
        title="Select Invoice",
        filetypes=[
            ("PDF files", "*.pdf"),
            ("Image files", "*.png *.jpg *.jpeg"),
            ("Excel files", "*.xlsx")
        ]
    )

    if file_path:
        selected_invoice_file = file_path
        invoice_label.config(text=file_path.split("/")[-1])


def get_file_type(file_path):
    file_path = file_path.lower()

    if file_path.endswith(".pdf"):
        return "pdf"

    if file_path.endswith((".png", ".jpg", ".jpeg")):
        return "image"

    if file_path.endswith(".xlsx"):
        return "excel"

    return "unknown"


def verify_documents():

    if not selected_po_file:
        messagebox.showwarning("Missing File", "Please select a PO file.")
        return

    if not selected_receipt_file:
        messagebox.showwarning("Missing File", "Please select a Receipt file.")
        return

    if not selected_invoice_file:
        messagebox.showwarning("Missing File", "Please select an Invoice file.")
        return

    po_type = get_file_type(selected_po_file)
    receipt_type = get_file_type(selected_receipt_file)
    invoice_type = get_file_type(selected_invoice_file)

    # PO
    if po_type == "pdf":
        po_data = read_po_pdf(selected_po_file)
    elif po_type == "excel":    
        po_data = read_po_excel(selected_po_file)
    elif po_type == "image":
        po_data = read_po_image(selected_po_file)    
    else:
        messagebox.showwarning(
            "Unsupported File",
            "Umsupported PO file type."
        )
        return

    # Receipt
    if receipt_type == "pdf":
        receipt_data = read_receipt_pdf(selected_receipt_file)
    elif receipt_type == "excel":
        receipt_data = read_receipt_excel(selected_receipt_file)
    elif receipt_type == "image":
        receipt_data = read_receipt_image(selected_receipt_file)    
    else:
        messagebox.showwarning(
            "Unsupported File",
            "Unsupported Receipt file type."
        )  
        return  

    # Invoice
    if invoice_type == "pdf":
        invoice_data = read_invoice_pdf(selected_invoice_file)
    elif invoice_type == "excel":
        invoice_data = read_invoice_excel(selected_invoice_file)
    elif invoice_type == "image":
        invoice_data = read_invoice_image(selected_invoice_file)    
    else:
        messagebox.showwarning(
            "Unsupported File",
            "Unsupported Invoice file type."
        )  
        return  
    
    po, receipt, invoice = normalize_all(
        po_data,
        receipt_data,
        invoice_data 
    )

    issues_found = 0

    issues_found += check_quantity(
        po["quantity"],
        receipt["quantity"],
        invoice["quantity"]
    )

    issues_found += check_price(
        po["price"],
        invoice["price"]
    )

    issues_found += check_po_number(
        po["number"],
        invoice["po_number"]
    )

    issues_found += check_supplier(
        po["supplier"],
        invoice["supplier"]
    )

    issues_found += check_vat(
        invoice["subtotal"],
        invoice["vat"]
    )

    issues_found += check_total(
        invoice["subtotal"],
        invoice["vat"],
        invoice["total"]
    )

    arabic_explanation, english_explanation = explain_result(
        po,
        receipt,
        invoice,
        issues_found
    )

    if issues_found == 0:
        dashboard_result = "FINAL RESULT: MATCHED"
    else:
        dashboard_result = "FINAL RESULT: REVIEW NEEDED"

    show_dashboard(
        dashboard_result,
        arabic_explanation,
        english_explanation,
        po,
        receipt,
        invoice
    )


def show_uplead_window():
    global po_label, receipt_label, invoice_label

    window = tk.Tk()
    window.title("Upload Documents")
    window.geometry("500x400")

    title = tk.Label(
        window,
        text="UPload Documents",
        font=("Arial", 18, "bold")
    )
    title.pack(pady=20)

    po_button = tk.Button(
        window,
        text="Select PO",
        command=select_po_file
    )
    po_button.pack(pady=10)

    po_label = tk.Label(
        window,
        text="NO PO selected",
        font=("Arial", 10)
    )
    po_label.pack()

    receipt_button = tk.Button(
        window,
        text="Select Receipt",
        command=select_receipt_file
    )
    receipt_button.pack(pady=10)

    receipt_label = tk.Label(
        window,
        text="NO Receipt selected",
        font=("Arial", 10)
    )
    receipt_label.pack()

    invoice_button = tk.Button(
        window,
        text="Select Invoice",
        command=select_invoice_file
    )
    invoice_button.pack(pady=10)

    invoice_label = tk.Label(
        window,
        text="NO Invoice selected",
        font=("Arial", 10)
    )
    invoice_label.pack()

    verify_button = tk.Button(
        window,
        text="Verify Documents",
        command=verify_documents,
        font=("Arial", 12 , "bold")
    )
    verify_button.pack(pady=25)

    window.mainloop()


def export_report(result_text):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
    )

    if not file_path:
        return

    with open(file_path, "w", encoding="utf-8") as file:
        file.write(result_text)

    messagebox.showinfo("Saved", "Report saved successfully.") 
   

def show_dashboard(
    result_text, 
    arabic_explanation, 
    english_explanation,
    po, 
    receipt, 
    invoice
):
    window = tk.Tk()
    window.title("Invoice Verification Dashboard")
    window.geometry("600x600")

    title = tk.Label(
        window,
        text="Invoice Verification Dashboard",
        font=("Arial", 18, "bold")
    )
    title.pack(pady=20)

    result_label = tk.Label(
        window,
        text=result_text,
        font=("Arial", 16, "bold")
    )
    result_label.pack(pady=15)

    explanation_label = tk.Label(
        window,
        text=arabic_explanation,
        font=("Arial", 9),
        wraplength=420,
        justify="center"
    )

    explanation_label.pack(pady=10)

    details_text = (
        f"PO Number: {po['number']}\n"
        f"Supplier: {po['supplier']}\n"
        f"Ordered Quantity: {po['quantity']}\n"
        f"Received Quantity: {receipt['quantity']}\n"
        f"Invoiced Quantity: {invoice['quantity']}\n"
        f"Unit Price: {invoice['price']}\n"
        f"VAT: {invoice['vat']}\n"
        f"Total: {invoice['total']}"
    )

    details_label = tk.Label(
        window,
        text=details_text,
        font=("Arial", 10),
        justify="left"
    )

    details_label.pack(pady=10)

    report_text = (
        f"{result_text}\n\n"
        f"{arabic_explanation}\n\n"
        f"{details_text}"
    )

    export_button = tk.Button(
        window,
        text="Export Report",
        command=lambda: export_report(report_text),
        font=("Arial", 11, "bold"),
        width=18
    )

    export_button.pack(pady=15)

    window.mainloop()


if __name__ == "__main__":
    show_uplead_window()      
