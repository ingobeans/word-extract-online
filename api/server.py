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

def get_all_elements(element, elements = []):
    elements.append(element)

    for child in element:
        get_all_elements(child, elements)
    
    return elements

def extract_text_from_xml(xml:str, links)->str:
    

    root = ET.fromstring(xml)

    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    t = ""

    in_link = False
    for element in get_all_elements(root):
        if element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink":
            in_link = True
            id = element.get(r"{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            t += "<a href=\"" + links[id] + "\">"
        elif element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r":
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
            if in_link:
                in_link = False
                t += "</a>"

        elif element.tag == r"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing":
            t += "<br>"

    return t

def get_links(xml_data):
    root = ET.fromstring(xml_data)

    # Dictionary to store results
    hyperlink_dict = {}

    # Iterate through each Relationship element
    for relationship in get_all_elements(root):
        if relationship.get("Type") != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
            continue
        id_value = relationship.get('Id')
        target_value = relationship.get('Target')

        # Add to dictionary
        hyperlink_dict[id_value] = target_value

    print(hyperlink_dict)
    return hyperlink_dict

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

            with zip_ref.open('word/_rels/document.xml.rels') as xml_file:
                links_content = xml_file.read().decode('utf-8')
            
            with zip_ref.open('word/document.xml') as xml_file:
                document_content = xml_file.read().decode('utf-8')

        return jsonify({'document_content': extract_text_from_xml(document_content, get_links(links_content))})

    except Exception as e:
        return jsonify({'error': f'Error processing the file: {str(e)}'})