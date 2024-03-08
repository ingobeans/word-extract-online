from flask import Flask, request, jsonify, render_template
from zipfile import ZipFile
from html_sanitizer import Sanitizer
import xml.etree.ElementTree as ET
import io

sanitizer = Sanitizer()
app = Flask(__name__)

def create_span_html(text="", font="arial", size=12, bold=False, italics=False, underline=False):
    style_string = f'font-family: {font}; font-size: {size}pt;'
    if bold:
        style_string += ' font-weight: bold;'
    if italics:
        style_string += ' font-style: italic;'
    if underline:
        style_string += ' text-decoration: underline;'

    html_string = f'<span style="{style_string}">{text}</span>'
    return html_string

def extract_text_from_xml(xml:str)->str:
    def recursive_loop(element, elements = []):
        elements.append(element)

        for child in element:
            recursive_loop(child, elements)
        
        return elements

    root = ET.fromstring(xml)

    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    t = ""

    for element in recursive_loop(root):
        if element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r":
            text = ""
            font = "arial"
            size = 12
            bold = False
            italics = False
            underline = False
            for child in element:
                tag = child.tag.removeprefix(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}")
                if tag == "t":
                    text = child.text
                elif tag == "rPr":
                    for data_child in child:
                        if data_child.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts":
                            font = data_child.get(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii")
                        elif data_child.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz":
                            size = int(data_child.get(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")) / 2
                        elif data_child.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b":
                            bold = data_child.get(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "1"
                        elif data_child.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i":
                            italics = data_child.get(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "1"
                        elif data_child.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u":
                            underline = True
            
            t += create_span_html(text, font, size, bold, italics, underline)

        elif element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing":
            t += "<br>"

    return t

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

        return jsonify({'document_content': extract_text_from_xml(document_content)})

    except Exception as e:
        return jsonify({'error': f'Error processing the file: {str(e)}'})
