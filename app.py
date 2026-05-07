import os
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, jsonify

from utils.docx_parser import DocxParser
from utils.html_builder import HtmlBuilder
from utils.pdf_generator import PdfGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB limit

UPLOAD_FOLDER = '/tmp/docx_uploads'
OUTPUT_FOLDER = '/tmp/pdf_outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado.'}), 400

    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado.'}), 400

    if not file.filename.lower().endswith('.docx'):
        return jsonify({'error': 'Apenas arquivos DOCX são suportados.'}), 400

    file_id = str(uuid.uuid4())
    upload_path = os.path.join(UPLOAD_FOLDER, f'{file_id}.docx')
    output_path = os.path.join(OUTPUT_FOLDER, f'{file_id}.pdf')
    original_stem = os.path.splitext(file.filename)[0]

    file.save(upload_path)

    try:
        parser = DocxParser(upload_path)
        content = parser.parse()

        # Use filename as fallback title
        if not content['metadata'].get('title'):
            content['metadata']['title'] = original_stem

        builder = HtmlBuilder(content)
        html_content = builder.build()

        generator = PdfGenerator()
        generator.generate(html_content, output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=f'{original_stem}.pdf',
            mimetype='application/pdf',
        )

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        return jsonify({'error': str(exc), 'traceback': tb}), 500

    finally:
        if os.path.exists(upload_path):
            os.remove(upload_path)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
