import streamlit as st
import tempfile
import os

from app import (
    read_po_excel,
    read_receipt_excel,
    read_invoice_excel,
    normalize_all,
    check_quantity,
    check_price,
    check_po_number,
    check_supplier,
    check_vat,
    check_total,
    explain_result,
    read_po_pdf,
    read_receipt_pdf,
    read_invoice_pdf,
    read_po_image,
    read_receipt_image,
    read_invoice_image
)


# def check_password():
 #   password = st.text_input("Enter Password", type="password")

 #   if password == st.secrets["APP_PASSWORD"]:
  #      return True

 #   if password:
 #       st.error("Incorrect password")

 #   return False

 
 # if not check_password():
 #   st.stop()


MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

st.title("Invoice Verification")

st.info("Beta,Version: Please upload sample or non-sensitive documents only. Do not upload confidential financial data.")

po_file = st.file_uploader(
    "Upload PO",
    type=["pdf", "xlsx", "png", "jpg", "jpeg"]
)

receipt_file = st.file_uploader(
    "Upload Receipt",
    type=["pdf", "xlsx", "png", "jpg", "jpeg"]
)

invoice_file = st.file_uploader(
    "Upload Invoice",
    type=["pdf", "xlsx", "png", "jpg", "jpeg"]

)

uploaded_files = [po_file, receipt_file, invoice_file]

for uploaded_file in uploaded_files:
    if uploaded_file is not None and uploaded_file.size > MAX_FILE_SIZE:
        st.error(f"{uploaded_file.name} is too large. Maximum size is 10 MB.")
        st.stop()

if po_file and receipt_file and invoice_file:
    st.success("All three files uploaded successfully.")

    if st.button("Verify Documents"):
        with tempfile.TemporaryDirectory() as temp_dir:
            po_path = os.path.join(temp_dir, "po" + os.path.splitext(po_file.name)[1].lower()) 
            receipt_path = os.path.join(temp_dir, "receipt" + os.path.splitext(receipt_file.name)[1].lower())
            invoice_path = os.path.join(temp_dir, "invoice" + os.path.splitext(invoice_file.name)[1].lower())

            with open(po_path, "wb") as f:
                f.write(po_file.getbuffer())

            with open(receipt_path, "wb") as f:
                f.write(receipt_file.getbuffer())

            with open(invoice_path, "wb") as f:
                f.write(invoice_file.getbuffer())

            st.success("Files are ready for verification.") 


            try:
                if po_path.lower().endswith(".pdf"):
                    po_data = read_po_pdf(po_path)
                elif po_path.lower().endswith((".png", ".jpg", ".jpeg")):
                    po_data = read_po_image(po_path)
                else:
                    po_data = read_po_excel(po_path) 

                if receipt_path.lower().endswith(".pdf"):
                    receipt_data = read_receipt_pdf(receipt_path) 
                elif receipt_path.lower().endswith((".png", ".jpg", ".jpeg")):
                    receipt_data = read_receipt_image(receipt_path)
                else:
                    receipt_data = read_receipt_excel(receipt_path) 

                if invoice_path.lower().endswith(".pdf"):
                    invoice_data = read_invoice_pdf(invoice_path)
                elif invoice_path.lower().endswith((".png", ".jpg", ".jpeg")):
                    invoice_data = read_invoice_image(invoice_path)
                else:
                    invoice_data = read_invoice_excel(invoice_path)

                po, receipt, invoice = normalize_all(
                    po_data,
                    receipt_data,
                    invoice_data
                )                             

                if (
                    not po["number"]
                    or not po["supplier"]
                    or po["quantity"] is None
                    or receipt["quantity"] is None
                    or not invoice["po_number"]
                    or not invoice["supplier"]
                    or invoice["quantity"] is None
                    or invoice["price"] is None
                    or invoice["vat"] is None
                    or invoice["total"] is None    
                ):
                    raise ValueError("Invalid document data")
                
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

                if issues_found == 0:
                    st.success("FINAL RESULT: MATCHED")
                else:
                    st.error("FINAL RESULT: REVIEW NEEDED")

                    if po["quantity"] != receipt["quantity"]:
                        st.warning(
                            f"Quantity mismatch: Ordered {po['quantity']}, "
                            f"Received {receipt['quantity']}."
                        ) 

                st.subheader("Saudi Invoice Checks") 

                if invoice["vat"] is not None and invoice["subtotal"] is not None:
                    expected_vat = invoice["subtotal"] * 0.15

                    if abs(invoice["vat"] - expected_vat) < 0.01:
                        st.success("VAT calculation is consistent.")
                    else:
                        st.warning("VAT calculation may be inconsistent.") 
                else:
                    st.warning("VAT information is missing.")

                if (
                    invoice["po_number"]
                    and invoice["supplier"]
                    and invoice["total"] is not None
                ):
                    st.success("Basic invoice information is present.")
                else:
                    st.warning("Some basic invoice information is missing.")
                        


                st.write("PO Number:", po["number"]) 
                st.write("Supplier:", po["supplier"]) 
                st.write("Ordered Quantity:", po["quantity"]) 
                st.write("Received Quantity:", receipt["quantity"]) 
                st.write("Invoice Quantity:", invoice["quantity"])
                st.write("Unit Price:", invoice["price"])
                st.write("vat:", invoice["vat"])
                st.write("Total:", invoice["total"])

                report_text = f"""
            FINAL RESULT: {"MATCHED" if issues_found == 0 else "REVIEW NEEDED"}

            PO Number: {po["number"]}
            Supplier: {po["supplier"]}
            Ordered Quantity: {po["quantity"]}
            Received Quantity: {receipt["quantity"]}
            Invoice Quantity: {invoice["quantity"]}
            Unit Price: {invoice["price"]}
            VAT: {invoice["vat"]}
            Total: {invoice["total"]}
            """

                st.download_button(
                    label="Download Report",
                    data=report_text,
                    file_name="invoice_report_text",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"Error details: {e}")
                st.error(
            ".حدث خطأ أثناء قراءة أو التحقق من الملفات"
            ".يرجى التأكد من أن الملفات صحيحة ومدعومة"
                )
                
                    

