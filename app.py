# Complete app.py (replace your existing file)

```python
import os
import json
import base64
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template, request, jsonify
from docx import Document
from openai import OpenAI

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-in-production")

# ---------------------------
# DEEPSEEK API CONFIGURATION
# ---------------------------
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    print("⚠️ DEEPSEEK_API_KEY not set")

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com/v1"
)

CV_PATH = "Master_CV.docx"
COVER_PATH = "Cover_Template.docx"

# ---------------------------
# DEEPSEEK CALL
# ---------------------------
def call_deepseek(prompt):
    """Call DeepSeek API and return JSON"""
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "You are an expert CV tailoring assistant. Return ONLY valid JSON."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.4,
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

# ---------------------------
# READ DOCX
# ---------------------------
def read_docx(file_path):
    """Extract text from a .docx file"""
    doc = Document(file_path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)

# ---------------------------
# TAILOR CV INCLUDING EXPERIENCE
# ---------------------------
def tailor_cv_deep(cv_text, job_description):
    """Tailor summary, skills, and experience"""

    prompt = f"""
You are a professional CV tailoring expert.

MASTER CV:
{cv_text}

JOB DESCRIPTION:
{job_description}

TASK:
Rewrite the CV so it aligns as closely as possible with the job description,
while remaining 100% truthful to the candidate's actual experience.

RULES:
1. Rewrite the professional summary to match the job requirements.
2. Rewrite the skills section to emphasize relevant competencies.
3. Rewrite each experience bullet point to highlight the most relevant achievements and responsibilities.
4. Do NOT invent new jobs, achievements, numbers, or responsibilities.
5. Keep the same job titles, employers, and dates.
6. Return ONLY valid JSON.

Return JSON in this exact structure:
{{
    "tailored_summary": "...",
    "tailored_skills": ["Skill 1", "Skill 2"],
    "tailored_experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "Rewritten bullet 1",
            "Rewritten bullet 2"
        ],
        "Senior Advisor – Projects Planning & Grants Development | Aug 2021 – Jan 2023": [
            "Rewritten bullet 1"
        ]
    }}
}}
"""

    return call_deepseek(prompt)

# ---------------------------
# TAILOR COVER LETTER
# ---------------------------
def tailor_cover_letter_deep(cover_text, cv_text, job_description):
    """Generate tailored cover letter"""

    prompt = f"""
You are a professional cover letter writer.

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

Create a 3-paragraph cover letter that perfectly matches this job.
Return ONLY the cover letter text.
"""

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "You are an expert cover letter writer. Return only the cover letter text."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.4
    )

    return response.choices[0].message.content

# ---------------------------
# UPDATE SUMMARY
# ---------------------------
def update_summary(doc, new_summary):
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip().lower() == "summary":
            if i + 1 < len(doc.paragraphs):
                p = doc.paragraphs[i + 1]
                if p.runs:
                    p.runs[0].text = new_summary
                    for run in p.runs[1:]:
                        run.text = ""
                else:
                    p.text = new_summary
            break

# ---------------------------
# UPDATE SKILLS
# ---------------------------
def update_skills(doc, new_skills):
    for i, para in enumerate(doc.paragraphs):
        if "skill" in para.text.strip().lower():
            j = i + 1

            # Clear old skills until Experience section
            while j < len(doc.paragraphs):
                text = doc.paragraphs[j].text.strip().lower()
                if "experience" in text:
                    break
                doc.paragraphs[j].text = ""
                j += 1

            # Insert new skills
            for idx, skill in enumerate(new_skills):
                if i + 1 + idx < len(doc.paragraphs):
                    doc.paragraphs[i + 1 + idx].text = f"• {skill}"
            break

# ---------------------------
# UPDATE EXPERIENCE
# ---------------------------
def update_experience(doc, tailored_experience):
    """Update experience bullets while preserving job titles and dates"""

    for i, para in enumerate(doc.paragraphs):
        job_title = para.text.strip()

        if job_title in tailored_experience:
            new_bullets = tailored_experience[job_title]

            # Find the employer line (usually next paragraph)
            employer_index = i + 1

            # Find the next section/job title
            j = employer_index + 1
            while j < len(doc.paragraphs):
                text = doc.paragraphs[j].text.strip()

                # Stop at Education or another job title
                if text in tailored_experience.keys() or text.lower() == "education":
                    break
                j += 1

            # Clear old bullets
            for k in range(employer_index + 1, j):
                doc.paragraphs[k].text = ""

            # Insert new bullets
            for idx, bullet in enumerate(new_bullets):
                target_index = employer_index + 1 + idx
                if target_index < len(doc.paragraphs):
                    doc.paragraphs[target_index].text = f"• {bullet}"

# ---------------------------
# PROCESS APPLICATION
# ---------------------------
def process_application(job_description):
    if not os.path.exists(CV_PATH):
        return None, None, f"CV file not found: {CV_PATH}"

    if not os.path.exists(COVER_PATH):
        return None, None, f"Cover letter template not found: {COVER_PATH}"

    cv_text = read_docx(CV_PATH)
    cover_text = read_docx(COVER_PATH)

    result = tailor_cv_deep(cv_text, job_description)

    tailored_summary = result.get("tailored_summary", "")
    tailored_skills = result.get("tailored_skills", [])
    tailored_experience = result.get("tailored_experience", {})

    # Update CV
    doc = Document(CV_PATH)
    update_summary(doc, tailored_summary)
    update_skills(doc, tailored_skills)
    update_experience(doc, tailored_experience)

    cv_output = BytesIO()
    doc.save(cv_output)
    cv_output.seek(0)

    # Update Cover Letter
    tailored_cover = tailor_cover_letter_deep(cover_text, cv_text, job_description)

    cover_doc = Document(COVER_PATH)
    paragraphs = [p for p in tailored_cover.split("\n") if p.strip()]

    for i, para in enumerate(cover_doc.paragraphs):
        if i < len(paragraphs):
            para.text = paragraphs[i]

    cover_output = BytesIO()
    cover_doc.save(cover_output)
    cover_output.seek(0)

    return cv_output, cover_output, "Success"

# ---------------------------
# ROUTES
# ---------------------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/tailor', methods=['POST'])
def tailor():
    job_description = request.form.get('job_description', '').strip()

    if not job_description:
        return jsonify({'error': 'Please paste a job description'}), 400

    cv_output, cover_output, message = process_application(job_description)

    if cv_output and cover_output:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return jsonify({
            'success': True,
            'cv_filename': f'Tailored_CV_{timestamp}.docx',
            'cover_filename': f'Tailored_Cover_{timestamp}.docx',
            'cv_data': base64.b64encode(cv_output.getvalue()).decode('utf-8'),
            'cover_data': base64.b64encode(cover_output.getvalue()).decode('utf-8')
        })

    return jsonify({'error': message}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy'})

# ---------------------------
# RUN APP
# ---------------------------
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)
```
