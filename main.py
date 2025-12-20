from flask import Flask, request, send_file, jsonify
import logging
import tempfile
import zipfile
from pathlib import Path
from lxml import etree
import shutil
import os

# ---------------- LOGGING ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

app = Flask(__name__)

# ---------------- CONFIG ----------------
BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates" / "FrontPage_Template.docx"

# Word XML namespace
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}

# ---------------- CORE LOGIC ----------------
def replace_placeholders_in_docx(template_path, replacements):
    """
    Replaces placeholders EVERYWHERE including shapes/textboxes
    by directly editing DOCX XML (DrawingML safe)
    """

    work_dir = Path(tempfile.mkdtemp())
    output_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    output_docx.close()

    try:
        # 1️⃣ Unzip DOCX
        with zipfile.ZipFile(template_path, "r") as zip_ref:
            zip_ref.extractall(work_dir)

        doc_xml = work_dir / "word" / "document.xml"

        if not doc_xml.exists():
            raise RuntimeError("document.xml not found in DOCX")

        # 2️⃣ Parse XML
        tree = etree.parse(str(doc_xml))
        root = tree.getroot()

        # 3️⃣ Replace text nodes (includes shapes)
        for text_node in root.xpath("//w:t", namespaces=NS):
            if text_node.text:
                original = text_node.text
                for key, value in replacements.items():
                    if key in original:
                        original = original.replace(key, value)
                text_node.text = original

        # 4️⃣ Save modified XML
        tree.write(
            str(doc_xml),
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes"
        )

        # 5️⃣ Zip back to DOCX
        with zipfile.ZipFile(output_docx.name, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for file in work_dir.rglob("*"):
                zip_out.write(file, file.relative_to(work_dir))

        return output_docx.name

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


# ---------------- ROUTES ----------------
@app.route("/", methods=["GET"])
def health():
    logging.info("❤️ Health check hit")
    return "OK", 200


@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        logging.info("📥 /generate called")

        data = request.get_json(force=True)
        logging.info(f"📦 Payload received: {data}")

        if not TEMPLATE_PATH.exists():
            logging.error("❌ Template DOCX missing")
            return jsonify(error="Template DOCX not found"), 500

        # STRICT replacement map
        replacements = {
            "{{TEST_NAME}}": data.get("test", ""),
            "{{CLASS}}": data.get("class", ""),
            "{{PHASE}}": data.get("phase", ""),
            "{{SET}}": data.get("set", ""),
            "{{DATE}}": data.get("date", ""),
            "{{Physics}}": data.get("physics", ""),
            "{{ROI/KPM}}": data.get("chemistry", ""),
            "{{Biology}}": data.get("biology", "")
        }

        logging.info("🧬 Replacing placeholders inside shapes (XML)")
        output_path = replace_placeholders_in_docx(
            TEMPLATE_PATH,
            replacements
        )

        logging.info(f"✅ DOCX generated at {output_path}")

        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"FrontPage_{data.get('set','SET')}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        logging.exception("❌ DOCX generation failed")
        return jsonify(error=str(e)), 500
