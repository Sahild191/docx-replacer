from flask import Flask, request, send_file
from docx import Document
import tempfile
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def health():
    return "ok"

@app.route("/", methods=["POST"])
def replace_placeholders():

    if "file" not in request.files:
        return "Missing file field", 400

    uploaded = request.files["file"]

    if uploaded.filename == "":
        return "Empty filename", 400

    # Temp paths
    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, "input.docx")
    output_path = os.path.join(temp_dir, "output.docx")

    # Save file
    uploaded.save(input_path)

    if not os.path.exists(input_path):
        return "File save failed", 500

    # Metadata
    replacements = {
        "{{CLASS}}": request.form.get("class", ""),
        "{{SET}}": request.form.get("set", ""),
        "{{TEST_NAME}}": request.form.get("test", ""),
        "{{PHASE}}": request.form.get("phase", ""),
        "{{DATE}}": request.form.get("date", "")
    }

    # Open DOCX safely
    doc = Document(input_path)

    def replace_paragraphs(paragraphs):
        for p in paragraphs:
            for k, v in replacements.items():
                if k in p.text:
                    p.text = p.text.replace(k, v)

    replace_paragraphs(doc.paragraphs)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_paragraphs(cell.paragraphs)

    doc.save(output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="FrontPage_Updated.docx"
    )
