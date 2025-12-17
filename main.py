from flask import Flask, request, send_file, jsonify
from docx import Document
import tempfile
import zipfile
import os

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        # 1️⃣ Get JSON data
        data = request.json
        zip_path = data.get("zip_path")
        values = data.get("values")

        if not zip_path or not values:
            return jsonify({"error": "Missing zip_path or values"}), 400

        # 2️⃣ Create temp directory
        temp_dir = tempfile.mkdtemp()

        zip_file_path = os.path.join(temp_dir, "template.zip")

        # 3️⃣ Download ZIP from Drive public URL
        import requests
        r = requests.get(zip_path)
        r.raise_for_status()

        with open(zip_file_path, "wb") as f:
            f.write(r.content)

        # 4️⃣ Extract ZIP
        with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        # 5️⃣ Find DOCX inside ZIP
        docx_path = None
        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith(".docx"):
                    docx_path = os.path.join(root, file)
                    break

        if not docx_path:
            return jsonify({"error": "No DOCX found in ZIP"}), 400

        # 6️⃣ Load DOCX
        doc = Document(docx_path)

        # 7️⃣ Replace placeholders
        for p in doc.paragraphs:
            for key, value in values.items():
                if key in p.text:
                    for run in p.runs:
                        run.text = run.text.replace(key, value)

        # 8️⃣ Save output
        output_path = os.path.join(temp_dir, "output.docx")
        doc.save(output_path)

        # 9️⃣ Return file
        return send_file(
            output_path,
            as_attachment=True,
            download_name="front_page.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def health():
    return "OK"
