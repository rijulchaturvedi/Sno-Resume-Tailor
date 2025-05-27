
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
from datetime import datetime
import io
from docx import Document
from docx.shared import Pt

app = Flask(__name__)
# CORS(app, origins=["chrome-extension://eejggmapnjhejendenjgekfeacdgcmki"])
CORS(app, origins=["*"])

@app.route("/tailor", methods=["POST", "OPTIONS"])
def tailor():
    if request.method == "OPTIONS":
        response = make_response()
        # response.headers["Access-Control-Allow-Origin"] = "chrome-extension://eejggmapnjhejendenjgekfeacdgcmki"
        response.headers["Access-Control-Allow-Origin"] = "*"
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    data = request.get_json()
    experience = data.get("experience", [])
    skills = data.get("skills", "")

    doc = Document("base_resume.docx")

    def replace_last_n_paragraphs(section_title, new_bullets, count, must_be_under=None):
        found_index = None
        for i, para in enumerate(doc.paragraphs):
            if section_title in para.text:
                if must_be_under:
                    context = [doc.paragraphs[j].text.strip().upper() for j in range(max(0, i - 3), i)]
                    if not any(must_be_under in line for line in context):
                        continue
                found_index = i
                break

        if found_index is None:
            print(f"❌ Section '{section_title}' not found.")
            return

        section_indices = []
        for j in range(found_index + 1, len(doc.paragraphs)):
            text = doc.paragraphs[j].text.strip()
            if len(text) > 0 and text.isupper():
                break
            if text:
                section_indices.append(j)

        if len(section_indices) < count:
            print(f"⚠️ Not enough paragraphs to replace under '{section_title}'.")
            return

        for k in range(count):
            idx = section_indices[-count + k]
            clean_bullet = new_bullets[k].replace("â€¢", "").replace("•", "").strip()
            doc.paragraphs[idx].text = clean_bullet
            for run in doc.paragraphs[idx].runs:
                run.font.size = Pt(10.5)
                run.font.name = "Times New Roman"

    replace_last_n_paragraphs("UNIVERSITY OF ILLINOIS URBANA-CHAMPAIGN", experience[0:2], 2, must_be_under="EXPERIENCE")
    replace_last_n_paragraphs("EXTUENT", experience[2:5], 3)
    replace_last_n_paragraphs("FRAPPE", experience[5:10], 5)

    for para in doc.paragraphs:
        if "Core Competencies" in para.text:
            if skills not in para.text:
                para.text = para.text.rstrip(" |") + " | " + skills
            break

    # Remove any lingering "• " anywhere in doc
    for para in doc.paragraphs:
        if "• " in para.text:
            para.text = para.text.replace("• ", "").strip()

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    filename = "Vrinda Menon Resume - " + datetime.now().strftime("%Y-%m-%d") + ".docx"
    response = make_response(send_file(output, as_attachment=True, download_name=filename))
    response.headers["Access-Control-Allow-Origin"] = "chrome-extension://eejggmapnjhejendenjgekfeacdgcmki"
    response.headers["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
