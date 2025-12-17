from flask import Flask, request, send_file
from docx import Document
import tempfile, zipfile, os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def health():
    return "OK", 200


@app.route("/generate", methods=["POST"])
def generate_docx():
    zip_file = request.files.get("zip")
    if not zip_file:
        return "ZIP missing", 400

    meta = request.form
    temp_dir = tempfile.mkdtemp()

    zip_path = os.path.join(temp_dir, "input.zip")
    zip_file.save(zip_path)

    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(temp_dir)

    docx_path = os.path.join(temp_dir, "FrontPage.docx")
    if not os.path.exists(docx_path):
        return "FrontPage.docx not found in ZIP", 400

    doc = Document(docx_path)

    replacements = {
        "{{CLASS}}": meta.get("class"),
        "{{SET}}": meta.get("set"),
        "{{TEST_NAME}}": meta.get("test"),
        "{{PHASE}}": meta.get("phase"),
        "{{DATE}}": meta.get("date")
    }

    def replace(container):
        for p in container.paragraphs:
            for k, v in replacements.items():
                if k in p.text:
                    p.text = p.text.replace(k, v)

    replace(doc)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace(cell)

    output_path = os.path.join(temp_dir, "Updated_FrontPage.docx")
    doc.save(output_path)

    return send_file(
        output_path,
        as_attachment=True,
        download_name="Updated_FrontPage.docx"
    )
