from flask import Flask, request, jsonify, send_file
from docx import Document
import base64
import tempfile
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def health():
    return "OK", 200


@app.route("/generate", methods=["POST"])
def generate_docx():
    try:
        # 1️⃣ Read JSON safely
        data = request.get_json(silent=True)
        if not data:
            return jsonify({"error": "Invalid or missing JSON body"}), 400

        # 2️⃣ Validate required fields
        required = ["file_base64", "class", "set", "test", "phase", "date"]
        missing = [k for k in required if k not in data]
        if missing:
            return jsonify({"error": f"Missing fields: {missing}"}), 400

        # 3️⃣ Decode base64 DOCX
        try:
            docx_bytes = base64.b64decode(data["file_base64"])
        except Exception as e:
            return jsonify({"error": f"Base64 decode failed: {str(e)}"}), 400

        # 4️⃣ Work in temp directory
        with tempfile.TemporaryDirectory() as tmp:
            input_path = os.path.join(tmp, "input.docx")
            output_path = os.path.join(tmp, "output.docx")

            # ✅ WRITE DOCX TO DISK (CRITICAL FIX)
            with open(input_path, "wb") as f:
                f.write(docx_bytes)

            # 🔒 Safety check
            if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
                return jsonify({"error": "DOCX file not written properly"}), 500

            # 5️⃣ Load DOCX
            doc = Document(input_path)

            # 6️⃣ Placeholder replacements
            replacements = {
                "{{CLASS}}": data["class"],
                "{{SET}}": data["set"],
                "{{TEST_NAME}}": data["test"],
                "{{PHASE}}": data["phase"],
                "{{DATE}}": data["date"]
            }

            # Replace in paragraphs
            for para in doc.paragraphs:
                for key, value in replacements.items():
                    if key in para.text:
                        para.text = para.text.replace(key, value)

            # Replace in tables (VERY IMPORTANT)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            for key, value in replacements.items():
                                if key in para.text:
                                    para.text = para.text.replace(key, value)

            # 7️⃣ Save output
            doc.save(output_path)

            # 8️⃣ Return DOCX
            return send_file(
                output_path,
                as_attachment=True,
                download_name="FrontPage.docx",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as e:
        # Absolute fallback
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
