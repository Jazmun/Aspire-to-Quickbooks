import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Invoice to Excel Converter", page_icon="📑", layout="wide")

st.title("📑 Landscaping Invoice to Excel Converter")
st.write("Upload your invoice PDF to extract all invoice items and download the formatted Excel spreadsheet.")

uploaded_file = st.file_uploader("Upload Invoice PDF", type=["pdf"])

def extract_invoice_data(pdf_bytes):
    records = []
    
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not text.strip():
                continue

            # 1. Invoice Number
            inv_match = re.search(r'Invoice\s+(\d+)', text)
            inv_num = inv_match.group(1) if inv_match else ""

            # 2. Invoice Date & Terms
            date_match = re.search(r'Date\s+PO#\s*\n\s*([0-9/]+)', text)
            inv_date_str = ""
            due_date_str = ""
            
            if date_match:
                raw_date = date_match.group(1).strip()
                try:
                    dt = datetime.strptime(raw_date, "%m/%d/%y")
                    inv_date_str = f"{dt.month}/{dt.day}/{dt.year}"
                    
                    if "Due on Receipt" in text:
                        due_date_str = f"{dt.month}/{dt.day:02d}/{dt.year}"
                    else: # Default Net 30
                        due_dt = dt + timedelta(days=30)
                        due_date_str = f"{due_dt.month}/{due_dt.day:02d}/{due_dt.year}"
                except:
                    inv_date_str = raw_date
                    due_date_str = raw_date

            # 3. Customer (Line 1 of Property Address)
            cust = ""
            prop_match = re.search(r'Property Address\s*\n([^\n]+)', text)
            if prop_match:
                cust = prop_match.group(1).strip()

            # 4. Tax
            tax_match = re.search(r'Sales Tax\s*\n?\s*\$?([\d,]+\.\d{2})', text)
            has_tax = "No"
            if tax_match:
                tax_val = float(tax_match.group(1).replace(',', ''))
                if tax_val > 0:
                    has_tax = "Yes"

            # 5. Amount and Line Items
            # Check for multiple line items or subtotal
            # Find lines matching pattern: Description ... $Amount
            lines = text.split('\n')
            found_items = []
            
            for line in lines:
                amount_pattern = re.search(r'^(.*?)\s+\$([\d,]+\.\d{2})$', line.strip())
                if amount_pattern:
                    item_desc = amount_pattern.group(1).strip()
                    item_val = float(amount_pattern.group(2).replace(',', ''))
                    # Exclude summary rows
                    if not any(k in item_desc for k in ["Subtotal", "Total", "Balance Due", "Current", "Retainage"]):
                        if item_val > 0:
                            found_items.append((item_desc, item_val))

            if not found_items:
                sub_match = re.search(r'Subtotal\s*\n?\s*\$([\d,]+\.\d{2})', text)
                sub_val = float(sub_match.group(1).replace(',', '')) if sub_match else 0.0
                found_items = [(f"Invoice {inv_num} Services", sub_val)]

            for desc, amt in found_items:
                records.append({
                    "Post": "Yes",
                    "Invoice Date": inv_date_str,
                    "Due Date": due_date_str,
                    "Invoice Number": inv_num,
                    "Transaction Type": "Invoice",
                    "Customer": cust,
                    "Vendor": "",
                    "Currency Code": "",
                    "Products/Services": "Side Jobs",
                    "Description": desc,
                    "Qty": 1,
                    "Discount %": "",
                    "Unit Price": amt,
                    "Category": "",
                    "Location": "",
                    "Class": "",
                    "Tax": has_tax
                })

    return records

def create_excel(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice Import"
    
    headers = [
        "Post", "Invoice Date", "Due Date", "Invoice Number", "Transaction Type",
        "Customer", "Vendor", "Currency Code", "Products/Services", "Description",
        "Qty", "Discount %", "Unit Price", "Category", "Location", "Class", "Tax"
    ]
    ws.append(headers)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, r in enumerate(records, start=2):
        row_values = [r[h] for h in headers]
        ws.append(row_values)
        fill_color = "F9FAFC" if row_idx % 2 == 0 else "FFFFFF"
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        for col_idx, cell in enumerate(ws[row_idx], start=1):
            cell.fill = row_fill
            cell.border = thin_border
            cell.font = Font(name="Calibri", size=10)
            col_name = headers[col_idx - 1]
            if col_name in ["Post", "Transaction Type", "Qty", "Tax", "Invoice Date", "Due Date", "Invoice Number"]:
                cell.alignment = Alignment(horizontal="center", vertical="top")
            elif col_name == "Unit Price":
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="top")
            elif col_name == "Description":
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, 40), 12)
    ws.column_dimensions['J'].width = 45

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if uploaded_file is not None:
    with st.spinner("Processing PDF and extracting line items..."):
        data = extract_invoice_data(uploaded_file.read())
        
    if data:
        st.success(f"Successfully processed {len(data)} invoice line items!")
        excel_data = create_excel(data)
        
        st.download_button(
            label="📥 Download Excel Spreadsheet",
            data=excel_data,
            file_name="Extracted_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("No invoice data found in the uploaded file.")
