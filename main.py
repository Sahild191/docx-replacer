from flask import Flask, request, send_file, jsonify
from docx import Document
import tempfile, os, base64

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate_docx():
    data = request.get_json()

    if not data or "file_base64" not in data:
        return jsonify({"error": "Missing file_base64"}), 400

    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, "input.docx")
    output_path = os.path.join(temp_dir, "output.docx")

    # Decode DOCX
    with open(input_path, "wb") as f:
        f.write(base64.b64decode(data["file_base64"]))

    doc = Document(input_path)

    replacements = {
        "{{CLASS}}": data.get("class", ""),
        "{{SET}}": data.get("set", ""),
        "{{TEST_NAME}}": data.get("test", ""),
        "{{PHASE}}": data.get("phase", ""),
        "{{DATE}}": data.get("date", "")
    }

    # Paragraphs
    for p in doc.paragraphs:
        for k, v in replacements.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    # Tables
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
        download_name="FrontPage.docx"
    )
