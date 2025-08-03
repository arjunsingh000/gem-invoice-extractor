from flask import Flask, render_template, request, send_file
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

import logging
logging.basicConfig(level=logging.DEBUG)


app = Flask(__name__)

def clean_excel_value(val):
    if isinstance(val, str):
        return re.sub(r"[\x00-\x1F]+", " ", val).strip()
    return val

def extract_field(pattern, text, group=1, default=""):
    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(group).strip() if match else default

def extract_multiline_block(start_label, end_labels, lines):
    start_index = None
    for i, line in enumerate(lines):
        if start_label.lower() in line.lower():
            start_index = i
            break

    if start_index is None:
        return ""

    block = []
    for line in lines[start_index + 1:]:
        if any(end.lower() in line.lower() for end in end_labels):
            break
        block.append(line.strip())

    return "\n".join([l for l in block if l])

def extract_from_pdfs(files):
    data = []

    for file in files:
        if file.filename.endswith(".pdf"):
            file.stream.seek(0)
            doc = fitz.open(stream=file.stream.read(), filetype="pdf")

            full_text = "\n".join(page.get_text() for page in doc)
            lines = full_text.splitlines()

            # Multiline block: Organisation and Buyer
            organisation_block = extract_multiline_block(
                start_label="Organisation Details",
                end_labels=["Buyer Details", "खरीदार विवरण", "Financial Approval"],
                lines=lines
            )
            buyer_block = extract_multiline_block(
                start_label="Buyer Details",
                end_labels=["Financial Approval", "Seller Details", "विक्रेता विवरण"],
                lines=lines
            )
            org_buyer_details = f"संस्‍थान विवरण:\n{organisation_block.strip()}\n\nखरीदार विवरण:\n{buyer_block.strip()}"

            # Seller Address block
            seller_address = extract_multiline_block(
                start_label="Address",
                end_labels=["Email ID", "GSTIN", "MSME", "Contact No", "Company Name"],
                lines=lines
            )
            seller_address = ", ".join(seller_address.split("\n"))

            record = {
                "File Name": file.filename,
                "Contract No": extract_field(r"Contract No[:\-]?\s*(GEMC-\d+)", full_text),
                "Generated Date": extract_field(r"Generated Date\s*:\s*(\d{1,2}-\w+-\d{4})", full_text),
                "Organisation & Buyer Details": org_buyer_details,
                "Seller Company Name": extract_field(r"Company Name\s*:\s*([^\n]*)", full_text),
                "Seller Phone": extract_field(r"Contact No\.?\s*:\s*-?(\d{10})", full_text),
                "Seller Email": extract_field(r"Email ID\s*:\s*([\w\.-]+@[\w\.-]+)", full_text),
                "Seller Address": seller_address,
                "Seller GSTIN": extract_field(r"GSTIN[:\s]*([A-Z0-9]+)", full_text),
                "Product Name": extract_field(r"Product Name\s*:\s*(.*?)\s*\|", full_text),
                "Brand": extract_field(r"Brand\s*:\s*(.*?)\s*\|", full_text),
                "Quantity": extract_field(r"(\d+)\s*pieces", full_text),
                "Unit Price": extract_field(r"pieces\s+([\d,]+)", full_text).replace(",", ""),
                "Total Price": extract_field(r"Total Order Value.*?(\d[\d,]*)", full_text).replace(",", ""),
                "Watts": extract_field(r"Rating\s*-\s*(\d+)\s*Watt", full_text)
            }

            data.append(record)

    df = pd.DataFrame(data)
    df = df.map(clean_excel_value)

    # Save using openpyxl for styling
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoices')

        # Apply wrap text to Organisation & Buyer Details column
        workbook = writer.book
        worksheet = writer.sheets['Invoices']

        for row in worksheet.iter_rows(min_row=2, min_col=4, max_col=4):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')

    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    try:
        files = request.files.getlist('pdfs')
        excel_output = extract_from_pdfs(files)
        return send_file(excel_output, as_attachment=True, download_name='gem_invoice_data.xlsx')
    except Exception as e:
        logging.exception("Internal server error occurred.")
        return "Internal Server Error", 500


@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.error(f"Unhandled Exception: {e}", exc_info=True)
    return "Internal server error occurred.", 500


if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

