from flask import Flask, request, jsonify, render_template
from zipfile import ZipFile
from html_sanitizer import Sanitizer
import xml.etree.ElementTree as ET
import io

sanitizer = Sanitizer()
app = Flask(__name__)

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
    last_font = None
    current_font_tag = None

    for element in recursive_loop(root):
        tag = element.tag.removeprefix(r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}")
        if tag == "t":
            t += element.text.replace("<","&lt;").replace(">","&gt;")
        if tag == "spacing":
            t += "\n"
        if tag == "rFonts":
            font = element.get(r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii')

            font_text = "</span><span style=\"font-family: '"+font+"';"
            current_font_tag = font_text
        if tag == "sz" and current_font_tag != None:
            size = element.get(r'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            font_text = "font-size: " + str(int(size)//2) + "pt;\">"

            current_font_tag += font_text
            if last_font == current_font_tag:
                current_font_tag = None
                continue

            last_font = current_font_tag
            t += current_font_tag
            current_font_tag = None
    
    t = t.replace("\n","<br>")

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
