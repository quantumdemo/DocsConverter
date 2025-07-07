from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import subprocess
# Removed: import docx2pdf
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
# DOCX → PDF (Using Pandoc)
# ----------------------------------------------
def convert_docx_to_pdf(docx_path, output_dir):
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
    try:
        # Command to convert DOCX to PDF using Pandoc
        # Pandoc will infer input and output formats from file extensions.
        cmd = [
            'pandoc', docx_path,
            '-o', pdf_path
        ]
        # Run the Pandoc command
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)

        if result.returncode != 0:
            # Pandoc failed, raise an exception with details from stderr
            error_message = result.stderr or "Pandoc failed with an unknown error."
            # Log the error for server-side debugging
            app.logger.error(f"Pandoc conversion error for {docx_path}: {error_message}")
            raise Exception(f"DOCX to PDF conversion failed: {error_message}")

        if not os.path.exists(pdf_path):
             # This case should ideally be caught by returncode != 0, but as a safeguard:
            app.logger.error(f"Pandoc conversion command seemed to succeed but output file {pdf_path} not found.")
            raise Exception(f"DOCX to PDF conversion failed: Output file not created by Pandoc.")

        return pdf_path
    except FileNotFoundError:
        # This would happen if pandoc command itself is not found (installation issue)
        app.logger.error("Pandoc command not found. Ensure Pandoc is installed and in PATH.")
        raise Exception("DOCX to PDF conversion failed: Pandoc command not found. Please check server configuration.")
    except Exception as e:
        # Catch any other exceptions during the process
        app.logger.error(f"An unexpected error occurred during DOCX to PDF conversion for {docx_path}: {e}")
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
