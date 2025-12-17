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


def replace_in_runs(paragraph, replacements):
    full_text = "".join(run.text for run in paragraph.runs)

    for key, value in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, value)

    # clear existing runs
    for run in paragraph.runs:
        run.text = ""

    # write back as single run (keeps paragraph formatting)
    paragraph.runs[0].text = full_text


@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Template DOCX not found"}), 500

        data = request.get_json(force=True)

        replacements = {
            "{{TEST_NAME}}": data["test"],
            "{{CLASS}}": data["class"],
            "{{PHASE}}": data["phase"],
            "{{SET}}": data["set"],
            "{{DATE}}": data["date"]
        }

        doc = Document(TEMPLATE_PATH)

        # 🔹 Replace in body
        for p in doc.paragraphs:
            replace_in_runs(p, replacements)

        # 🔹 Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_runs(p, replacements)

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
