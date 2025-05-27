
from flask import Flask, request, send_file, make_response
from flask_cors import CORS
from datetime import datetime
import io
from docx import Document

app = Flask(__name__)
CORS(app, origins=["chrome-extension://eejggmapnjhejendenjgekfeacdgcmki"])

@app.route("/tailor", methods=["POST", "OPTIONS"])
def tailor():
    if request.method == "OPTIONS":
        response = make_response()
        response.headers["Access-Control-Allow-Origin"] = "chrome-extension://eejggmapnjhejendenjgekfeacdgcmki"
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    data = request.get_json()
    experience = data.get("experience", [])
    skills = data.get("skills", "")

    doc = Document("base_resume.docx")

    def replace_bullets(section_title, new_bullets):
        for i in range(len(doc.paragraphs)):
            if section_title in doc.paragraphs[i].text:
                j = i + 1
                while j < len(doc.paragraphs) and (
                    doc.paragraphs[j].text.strip().startswith("•") or doc.paragraphs[j].text.strip() == ""
                ):
                    del doc.paragraphs[j]
                for idx, bullet in enumerate(new_bullets):
                    doc.paragraphs.insert(j + idx, doc.add_paragraph(bullet))
                break

    replace_bullets("UNIVERSITY OF ILLINOIS URBANA-CHAMPAIGN", experience[0:2])
    replace_bullets("EXTUENT", experience[2:5])
    replace_bullets("FRAPPE", experience[5:10])

    for para in doc.paragraphs:
        if "Core Competencies" in para.text:
            para.text = "Core Competencies - " + skills
            break

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
