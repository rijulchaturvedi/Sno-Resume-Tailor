
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

    def clean(text):
        return text.strip().lower().replace("\t", " ").replace("|", " ").replace("-", " ").replace("–", " ")

    def locate_role(company, role, bullets):
        for i in range(len(doc.paragraphs) - 1):
            if company.lower() in clean(doc.paragraphs[i].text) and i > 0 and "experience" in doc.paragraphs[i - 1].text.lower():
                for j in range(i + 1, len(doc.paragraphs)):
                    if role.lower() in clean(doc.paragraphs[j].text):
                        start = j + 1
                        end = start
                        while end < len(doc.paragraphs) and (
                            doc.paragraphs[end].text.strip().startswith("•") or doc.paragraphs[end].text.strip() == ""
                        ):
                            end += 1
                        for _ in range(end - start):
                            del doc.paragraphs[start]
                        for k, b in enumerate(bullets):
                            doc.paragraphs.insert(start + k, doc.add_paragraph(b))
                        return

    locate_role("UNIVERSITY OF ILLINOIS URBANA-CHAMPAIGN", "Product Data Analyst", experience[0:2])
    locate_role("EXTUENT", "Product Manager", experience[2:5])
    locate_role("FRAPPE", "Project Manager", experience[5:10])

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
