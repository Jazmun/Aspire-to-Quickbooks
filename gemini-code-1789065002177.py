import streamlit as st
import pdfplumber
import openpyxl
from io import BytesIO
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Invoice PDF to Excel Converter", layout="wide")
st.title("📄 Landscaping Invoice to Excel Converter")
st.write("Upload an invoice PDF to extract line items and export directly into your accounting import spreadsheet.")

uploaded_file = st.file_uploader("Choose an Invoice PDF", type=["pdf"])

def parse_invoices(file_bytes):
    rows = []
    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 1. Invoice Number
            inv_match = re.search(r'Invoice\s+(\d+)', text)
            inv_no = inv_match.group(1) if inv_match else ""
            
            # 2. Invoice Date & Terms
            date_match = re.search(r'Date\s+PO#\s*\n\s*([0-9/]+)', text)
            inv_date_raw = date_match.group(1) if date_match else ""
            
            # Parse Date & Terms for Due Date
            due_date_str = ""
            inv_date_str = ""
            if inv_date_raw:
                try:
                    dt = datetime.strptime(inv_date_raw, "%m/%d/%y")
                    inv_date_str = f"{dt.month}/{dt.day}/{dt.year}"
                    if "Due on Receipt" in text:
                        due_date_str = inv_date_str
                    else: # Net 30 default
                        due_dt = dt + timedelta(days=30)
                        due_date_str = f"{due_dt.month}/{due_dt.day:02d}/{due_dt.year}"
                except:
                    inv_date_str = inv_date_raw
                    due_date_str = inv_date_raw

            # 3. Customer (First line of Property Address)
            cust = ""
            addr_match = re.search(r'Property Address\s*\n([^\n]+)', text)
            if addr_match:
                cust = addr_match.group(1).strip()
                
            # 4. Tax check
            tax_match = re.search(r'Sales Tax\s*\n?\s*(\$[\d,]+\.\d{2})', text)
            has_tax = "Yes" if tax_match and tax_match.group(1) != "$0.00" else "No"
            
            # 5. Amount & Description
            # Capture line items or subtotal
            amount_match = re.search(r'Subtotal\s*\n?\s*\$([\d,]+\.\d{2})', text)
            amount = float(amount_match.group(1).replace(',', '')) if amount_match else 0.0
            
            # Add to output rows
            rows.append({
                "Post": "Yes",
                "Invoice Date": inv_date_str,
                "Due Date": due_date_str,
                "Invoice Number": inv_no,
                "Transaction Type": "Invoice",
                "Customer": cust,
                "Vendor": "",
                "Currency Code": "",
                "Products/Services": "Side Jobs",
                "Description": f"Invoice {inv_no} services",
                "Qty": 1,
                "Discount %": "",
                "Unit Price": amount,
                "Category": "",
                "Location": "",
                "Class": "",
                "Tax": has_tax
            })
    return rows

if uploaded_file is not None:
    data = parse_invoices(uploaded_file.read())
    st.success(f"Extracted {len(data)} invoice records!")
    
    # Generate Excel in memory
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = list(data[0].keys())
    ws.append(headers)
    for item in data:
        ws.append(list(item.values()))
        
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    st.download_button(
        label="📥 Download Excel Spreadsheet",
        data=out,
        file_name="Invoice_Import.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )