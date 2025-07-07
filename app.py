from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import subprocess
import docx2pdf # Added import
from pdf2docx import Converter
from PyPDF2 import PdfMerger
import atexit
import shutil

app = Flask(__name__)


UPLOAD = 'uploads'
OUT = 'converted'
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

# ----------------------------------------------
# DOCX → PDF (Using docx2pdf)
# ----------------------------------------------
def convert_docx_to_pdf(docx_path, output_dir):
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
    try:
        docx2pdf.convert(docx_path, pdf_path)
        return pdf_path
    except Exception as e:
        # It's good to catch a broad exception here, as docx2pdf can raise various errors
        # depending on the underlying issues (file corruption, missing dependencies it might still try to use, etc.)
        raise Exception(f"DOCX to PDF conversion failed: {e}")

# ----------------------------------------------
# Routes
# ----------------------------------------------
@app.route('/')
def home():
    return render_template('home.html')

@app.route('/')
def convert():
    return render_template('index.html')


@app.route('/convert/docx-to-pdf', methods=['POST'])
def docx_to_pdf():
    try:
        if 'file' not in request.files:
            return "No file uploaded", 400
            
        f = request.files['file']
        if f.filename == '':
            return "Empty filename", 400

        # Secure save input file
        fname = secure_filename(f.filename)
        in_path = os.path.join(UPLOAD, fname)
        f.save(in_path)

        # Convert and send result
        out_path = convert_docx_to_pdf(in_path, OUT)
        return send_file(out_path, as_attachment=True, download_name=f"{os.path.splitext(fname)[0]}.pdf")

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
        
        return send_file(out_path, as_attachment=True, download_name=f"{os.path.splitext(fname)[0]}.docx")
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
        return send_file(out_path, as_attachment=True, download_name="merged.pdf")
    except Exception as e:
        return f"Error merging PDFs: {str(e)}", 500

# ----------------------------------------------
# Cleanup (Critical for Render's ephemeral storage)
# ----------------------------------------------
@atexit.register
def cleanup():
    for folder in [UPLOAD, OUT]:
        if os.path.exists(folder):
            shutil.rmtree(folder)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
