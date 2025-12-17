from flask import Flask, request, send_file
from docx import Document
import tempfile
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def health():
    return "OK", 200


@app.route("/", methods=["POST"])
def replace_placeholders():

    file = request.files.get("file")
    if not file:
        return "No file received", 400

    meta = request.form

    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, "input.docx")
    output_path = os.path.join(temp_dir, "output.docx")

    file.save(input_path)

    doc = Document(input_path)

    replacements = {
        "{{CLASS}}": meta.get("class", ""),
        "{{SET}}": meta.get("set", ""),
        "{{TEST_NAME}}": meta.get("test", ""),
        "{{PHASE}}": meta.get("phase", ""),
        "{{DATE}}": meta.get("date", "")
    }

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)

    doc.save(output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="FrontPage_Updated.docx"
    )
