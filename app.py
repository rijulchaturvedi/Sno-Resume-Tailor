
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

    def locate_and_replace(company, role, new_bullets):
        for i in range(len(doc.paragraphs) - 1):
            para_text = doc.paragraphs[i].text.strip().upper()
            # Only proceed if this is in the EXPERIENCE section
            if company in doc.paragraphs[i].text.strip() and i > 0 and "EXPERIENCE" in doc.paragraphs[i - 1].text.upper():
                # Look for the role title immediately after
                for j in range(i + 1, len(doc.paragraphs) - 1):
                    if role.lower() in doc.paragraphs[j].text.strip().lower():
                        start = j + 1
                        end = start
                        while end < len(doc.paragraphs) and (
                            doc.paragraphs[end].text.strip().startswith("•") or doc.paragraphs[end].text.strip() == ""
                        ):
                            end += 1
                        for _ in range(end - start):
                            del doc.paragraphs[start]
                        for b, bullet in enumerate(new_bullets):
                            doc.paragraphs.insert(start + b, doc.add_paragraph(bullet))
                        return

    locate_and_replace("UNIVERSITY OF ILLINOIS URBANA-CHAMPAIGN", "Product Data Analyst - Research Assistant", experience[0:2])
    locate_and_replace("EXTUENT", "Product Manager", experience[2:5])
    locate_and_replace("FRAPPE", "Project Manager", experience[5:10])

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
