from flask import Flask, request, send_file, jsonify
from docx import Document
import os
import tempfile
import logging

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = "templates/FrontPage_Template.docx"

@app.route("/", methods=["GET"])
def health():
    return "OK", 200

@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        data = request.get_json()
        logging.info(f"📦 Payload received: {data}")

        if not os.path.exists(TEMPLATE_PATH):
            logging.error("❌ Template DOCX not found")
            return jsonify(error="Template DOCX not found"), 500

        doc = Document(TEMPLATE_PATH)

        replacements = {
            "{{TEST_NAME}}": data.get("test", ""),
            "{{CLASS}}": data.get("class", ""),
            "{{PHASE}}": data.get("phase", ""),
            "{{SET}}": data.get("set", ""),
            "{{DATE}}": data.get("date", "")
        }

        for para in doc.paragraphs:
            for key, val in replacements.items():
                if key in para.text:
                    para.text = para.text.replace(key, val)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            return send_file(
                tmp.name,
                as_attachment=True,
                download_name="FrontPage.docx"
            )

    except Exception as e:
        logging.exception("❌ DOCX generation failed")
        return jsonify(error=str(e)), 500
