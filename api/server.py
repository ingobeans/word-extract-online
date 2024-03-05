from flask import Flask, request, jsonify, render_template
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import io

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'})

    try:
        zip_content = io.BytesIO(file.read())
        with ZipFile(zip_content, 'r') as zip_ref:
            if 'word/document.xml' not in zip_ref.namelist():
                return jsonify({'error': 'No "word" folder found in the zip file'})

            with zip_ref.open('word/document.xml') as xml_file:
                document_content = xml_file.read().decode('utf-8')

            root = ET.fromstring(document_content)

            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            }

            text_elements = root.findall('.//w:t', namespaces=ns)

            t = ""
            for element in text_elements:
                raw_text = element.text
                if raw_text is not None:
                    t += raw_text

        return jsonify({'document_content': t})

    except Exception as e:
        return jsonify({'error': f'Error processing the file: {str(e)}'})
