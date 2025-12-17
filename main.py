from flask import Flask, request, send_file
from docx import Document
import tempfile
import os

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate_docx():

    if "file" not in request.files:
        return "No file received", 400

    file = request.files["file"]
    data = request.form

    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, "input.docx")
    output_path = os.path.join(temp_dir, "output.docx")

    file.save(input_path)

    doc = Document(input_path)

    replacements = {
        "{{CLASS}}": data.get("class", ""),
        "{{SET}}": data.get("set", ""),
        "{{TEST_NAME}}": data.get("test", ""),
        "{{PHASE}}": data.get("phase", ""),
        "{{DATE}}": data.get("date", "")
    }

    # Replace in paragraphs
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in replacements.items():
                        if k in p.text:
                            p.text = p.text.replace(k, v)

    doc.save(output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="FrontPage_Updated.docx"
    )
