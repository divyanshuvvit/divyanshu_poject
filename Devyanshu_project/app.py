
from flask import Flask, render_template, request, send_file
from io import BytesIO
import pandas as pd
from openpyxl import Workbook
import re
from docx import Document
from PyPDF2 import PdfReader

app = Flask(__name__)

# Function to extract email IDs and contact numbers from text
def extract_info(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'
    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    return ", ".join(emails), ", ".join(phones)

# Function to extract text from DOCX file
def extract_docx_text(file):
    doc = Document(file)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return "\n".join(full_text)

# Function to extract text from PDF file
def extract_pdf_text(file):
    pdf_reader = PdfReader(file)
    text = ''
    for page_num in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text()
    return text

# Main function to process CVs
def process_cvs(uploaded_files):
    extracted_data = []
    for file in uploaded_files:
        if file.mimetype == "application/pdf":
            # Process PDF files
            pdf_data = BytesIO(file.read())
            cv_text = extract_pdf_text(pdf_data)
        elif file.mimetype == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            # Process DOCX files
            docx_data = BytesIO(file.read())
            cv_text = extract_docx_text(docx_data)
        else:
            # Unsupported file format
            continue
        
        email, phone = extract_info(cv_text)
        extracted_data.append({"Email": email, "Phone": phone, "Text": cv_text})
    return extracted_data

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_files = request.files.getlist("file[]")
        if uploaded_files:
            cv_texts = process_cvs(uploaded_files)

            if cv_texts:
                df = pd.DataFrame(cv_texts)
                output_file = 'extracted_data.xls'
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return send_file(output_file, as_attachment=True)
            else:
                return "No valid CVs were processed."
    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)