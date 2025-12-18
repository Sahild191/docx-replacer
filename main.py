from flask import Flask, request, send_file, jsonify
import logging
import tempfile
import zipfile
from pathlib import Path
from lxml import etree
import shutil

# ---------------- LOGGING ----------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates" / "FrontPage_Template.docx"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


def replace_placeholders_in_docx(template_path, replacements):
    work_dir = Path(tempfile.mkdtemp())
    output_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    output_docx.close()

    try:
        with zipfile.ZipFile(template_path, "r") as zip_ref:
            zip_ref.extractall(work_dir)

        doc_xml = work_dir / "word" / "document.xml"
        tree = etree.parse(str(doc_xml))
        root = tree.getroot()

        for node in root.xpath("//w:t", namespaces=NS):
            if node.text:
                for k, v in replacements.items():
                    if k in node.text:
                        node.text = node.text.replace(k, v)

        tree.write(
            str(doc_xml),
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes"
        )

        with zipfile.ZipFile(output_docx.name, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for f in work_dir.rglob("*"):
                zip_out.write(f, f.relative_to(work_dir))

        return output_docx.name

    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/", methods=["GET"])
def health():
    return "OK", 200


@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        data = request.get_json(force=True)

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

        output_path = replace_placeholders_in_docx(
            TEMPLATE_PATH,
            replacements
        )

        return send_file(
            output_path,
            as_attachment=True,
            download_name=f"FrontPage_{data.get('set','SET')}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        logging.exception("DOCX generation failed")
        return jsonify(error=str(e)), 500
