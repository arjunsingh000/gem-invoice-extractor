from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import pdfplumber
import re
from io import BytesIO

app = Flask(__name__)

def clean_excel_value(val):
    if isinstance(val, str):
        return re.sub(r"[\x00-\x1F]+", " ", val).strip()
    return val

def extract_field(pattern, text, group=1, default=""):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(group).strip() if match else default

def extract_from_pdfs(files):
    data = []

    for file in files:
        if file.filename.endswith(".pdf"):
            with pdfplumber.open(file) as pdf:
                full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

            # Extract block of Organisation & Buyer info
            org_block = extract_field(r"(Organisation Details[\s\S]{50,400}?)\n\s*Buyer Details", full_text, group=1)
            buyer_block = extract_field(r"(Buyer Details[\s\S]{30,400}?)\n\s*\w", full_text, group=1)

            seller_address = extract_field(r"Address\s*:\s*(.*\n.*,\s*\w+,\s*\w+-\d+)", full_text)

            record = {
                "File Name": file.filename,
                "Contract No.": extract_field(r"Contract No[:\-]?\s*(GEMC-\d+)", full_text),
                "Generated Date": extract_field(r"Generated Date\s*:\s*(\d{1,2}-\w+-\d{4})", full_text),
                "Organisation & Buyer Details": (org_block + "\n\n" + buyer_block).strip(),
                "Seller Company": extract_field(r"Company Name\s*:\s*(.+)", full_text),
                "Seller Phone": extract_field(r"Contact No\.\s*:\s*([0-9\-]+)", full_text),
                "Seller Email": extract_field(r"Email ID\s*:\s*([\w\.\-@]+)", full_text),
                "Seller GSTIN": extract_field(r"GSTIN\s*:\s*([A-Z0-9]+)", full_text),
                "Seller Address": seller_address.replace("\n", ", ") if seller_address else "",
                "Product Name": extract_field(r"Product Name\s*:\s*(.+)", full_text),
                "Brand": extract_field(r"Brand\s*:\s*(.+)", full_text),
                "Quantity": extract_field(r"(\d+)\s+pieces", full_text),
                "Unit Price": extract_field(r"pieces\s+([\d,]+)", full_text),
                "Total Price": extract_field(r"Total Order Value.*?(\d[\d,]*)", full_text),
                "Wattage": extract_field(r"Rating\s*-\s*(\d+)\s*Watt", full_text)
            }

            data.append(record)

    df = pd.DataFrame(data)
    df = df.map(clean_excel_value)
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    files = request.files.getlist('pdfs')
    excel_output = extract_from_pdfs(files)
    return send_file(excel_output, as_attachment=True, download_name='gem_invoice_data.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
