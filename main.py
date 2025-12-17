from flask import Flask, request, send_file, jsonify
from docx import Document
import tempfile
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "FrontPage_Template.docx")


@app.route("/", methods=["GET"])
def health():
    return "OK", 200


@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({
                "error": "Template DOCX not found",
                "expected_path": TEMPLATE_PATH
            }), 500

        data = request.get_json(force=True)

        required = ["test", "class", "phase", "set", "date"]
        for k in required:
            if k not in data:
                return jsonify({"error": f"Missing field: {k}"}), 400

        doc = Document(TEMPLATE_PATH)

        replacements = {
            "{{TEST_NAME}}": data["test"],
            "{{CLASS}}": data["class"],
            "{{PHASE}}": data["phase"],
            "{{SET}}": data["set"],
            "{{DATE}}": data["date"],
        }

        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    for run in para.runs:
                        run.text = run.text.replace(key, val)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)

        return send_file(
            tmp.name,
            as_attachment=True,
            download_name="FrontPage.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
