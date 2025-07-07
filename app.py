from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import subprocess
from pdf2docx import Converter
from PyPDF2 import PdfMerger
from docx import Document

app = Flask(__name__)

UPLOAD = 'uploads'
OUT = 'converted'
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

# Function to convert DOCX to PDF using LibreOffice
def convert_docx_to_pdf(docx_path, output_dir):
    try:
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path
        ], check=True)
        pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
        return os.path.join(output_dir, pdf_filename)
    except subprocess.CalledProcessError:
        raise Exception("Failed to convert DOCX to PDF using LibreOffice.")

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/convert')
def convert():
    return render_template('index.html')

@app.route('/convert/docx-to-pdf', methods=['POST'])
def docx_to_pdf():
    try:
        f = request.files['file']
        fname = secure_filename(f.filename)
        in_path = os.path.join(UPLOAD, fname)
        f.save(in_path)

        out_path = convert_docx_to_pdf(in_path, OUT)
        return send_file(out_path, as_attachment=True, mimetype='application/pdf')
    except Exception as e:
        return f"Error converting DOCX to PDF: {str(e)}", 500

@app.route('/convert/pdf-to-docx', methods=['POST'])
def pdf_to_docx():
    try:
        f = request.files['file']
        fname = secure_filename(f.filename)
        in_path = os.path.join(UPLOAD, fname)
        out_path = os.path.join(OUT, fname.replace('.pdf', '.docx'))
        f.save(in_path)

        cv = Converter(in_path)
        cv.convert(out_path)
        cv.close()
        return send_file(out_path, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return f"Error converting PDF to DOCX: {str(e)}", 500

@app.route('/convert/merge-pdfs', methods=['POST'])
def merge_pdfs():
    try:
        files = request.files.getlist('files')
        merger = PdfMerger()
        out_path = os.path.join(OUT, 'merged.pdf')
        for f in files:
            fname = secure_filename(f.filename)
            path = os.path.join(UPLOAD, fname)
            f.save(path)
            merger.append(path)
        merger.write(out_path)
        merger.close()
        return send_file(out_path, as_attachment=True, mimetype='application/pdf')
    except Exception as e:
        return f"Error merging PDFs: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
