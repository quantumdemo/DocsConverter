from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
# import pypandoc
from docx2pdf import convert as docx2pdf_convert
from pdf2docx import Converter
from PyPDF2 import PdfMerger
from docx import Document
import easyocr


app = Flask(__name__)

UPLOAD = 'uploads'
OUT = 'converted'
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

reader = easyocr.Reader(['en'], gpu=True)

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
        out_path = os.path.join(OUT, fname.replace('.docx', '.pdf'))
        f.save(in_path)

        docx2pdf_convert(in_path, out_path)  # uses MS Word on Windows/macOS :contentReference[oaicite:2]{index=2}
        # pypandoc.convert_file(in_path, 'pdf', outputfile=out_path) #For mobile user
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

@app.route('/convert/image-to-docx', methods=['POST'])
def image_to_docx():
    try:
        f = request.files['file']
        fname = secure_filename(f.filename)
        in_path = os.path.join(UPLOAD, fname)
        out_path = os.path.join(OUT, fname.rsplit('.',1)[0] + '_ocr.docx')
        f.save(in_path)

        texts = reader.readtext(in_path, detail=0)
        doc = Document()
        for line in texts:
            doc.add_paragraph(line)
        doc.save(out_path)
        return send_file(out_path, as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return f"Error extracting text from image: {str(e)}", 500

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
    app.run(debug=False)