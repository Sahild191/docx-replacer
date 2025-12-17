from flask import Flask, request, jsonify, send_file
from docx import Document
import tempfile
import os
import logging

# ---------------- LOGGING SETUP ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)
# ------------------------------------------------

app = Flask(__name__)

TEMPLATE_PATH = "FrontPage_Template.docx"  # stored in Git repo


@app.route("/generate", methods=["POST"])
def generate_docx():
    logger.info("📥 /generate called")

    try:
        data = request.get_json(force=True)
        logger.info(f"📦 Payload received: {data}")

        # Validate input
        required = ["test", "class", "phase", "set", "date"]
        for key in required:
            if key not in data:
                logger.error(f"❌ Missing field: {key}")
                return jsonify({"error": f"Missing field: {key}"}), 400

        # Template check
        if not os.path.exists(TEMPLATE_PATH):
            logger.error("❌ Template DOCX not found in repo")
            return jsonify({"error": "Template DOCX not found"}), 500

        logger.info("📄 Loading DOCX template")
        doc = Document(TEMPLATE_PATH)

        replacements = {
            "{{TEST_NAME}}": data["test"],
            "{{CLASS}}": data["class"],
            "{{PHASE}}": data["phase"],
            "{{SET}}": data["set"],
            "{{DATE}}": data["date"],
        }

        def replace_in_paragraph(paragraph):
            original = paragraph.text
            updated = original

            for key, val in replacements.items():
                if key in updated:
                    updated = updated.replace(key, val)

            if updated != original:
                paragraph.clear()
                paragraph.add_run(updated)
                logger.info(f"✏ Replaced text: '{original}' → '{updated}'")

        # BODY
        logger.info("🧩 Replacing placeholders in body")
        for paragraph in doc.paragraphs:
            replace_in_paragraph(paragraph)

        # TABLES
        logger.info("🧩 Replacing placeholders in tables")
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_paragraph(paragraph)

        # Save output
        temp_dir = tempfile.mkdtemp()
        output_path = os.path.join(temp_dir, "FrontPage_Output.docx")

        doc.save(output_path)
        logger.info(f"💾 DOCX saved at {output_path}")

        if not os.path.exists(output_path):
            logger.error("❌ DOCX file not written properly")
            return jsonify({"error": "DOCX file not written properly"}), 500

        logger.info("📤 Sending updated DOCX back to Apps Script")
        return send_file(
            output_path,
            as_attachment=True,
            download_name="FrontPage.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        logger.exception("🔥 Unhandled exception in /generate")
        return jsonify({"error": str(e)}), 500


@app.route("/", methods=["GET"])
def health():
    logger.info("❤️ Health check hit")
    return "OK", 200
