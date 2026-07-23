import os
import json
import base64
from datetime import datetime
from io import BytesIO

from flask import Flask, render_template, request, jsonify
from docx import Document
from openai import OpenAI

app = Flask(**name**)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")

# DeepSeek API configuration

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")

client = OpenAI(
api_key=DEEPSEEK_API_KEY,
base_url="https://api.deepseek.com/v1"
)

CV_PATH = "Master_CV.docx"
COVER_PATH = "Cover_Template.docx"

def call_deepseek(prompt):
response = client.chat.completions.create(
model="deepseek-chat",
messages=[
{
"role": "system",
"content": "You are an expert CV tailoring assistant. Return ONLY valid JSON."
},
{
"role": "user",
"content": prompt
}
],
temperature=0.4,
response_format={"type": "json_object"}
)

```
return json.loads(response.choices[0].message.content)
```

def read_docx(file_path):
doc = Document(file_path)
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
return "\n".join(paragraphs)

def tailor_cv_deep(cv_text, job_description):
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
]
}}
}}
"""

```
return call_deepseek(prompt)
```

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
prompt = f"""
You are a professional cover letter writer.

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

Create a professional 3-paragraph cover letter that perfectly matches this job.
Return ONLY the cover letter text.
"""

```
response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {
            "role": "system",
            "content": "You are an expert cover letter writer. Return only the cover letter text."
        },
        {
            "role": "user",
            "content": prompt
        }
    ],
    temperature=0.4
)

return response.choices[0].message.content
```

def update_summary(doc, new_summary):
for i, para in enumerate(doc.paragraphs):
if para.text.strip().lower() == "summary":
if i + 1 < len(doc.paragraphs):
doc.paragraphs[i + 1].text = new_summary
break

def update_skills(doc, new_skills):
for i, para in enumerate(doc.paragraphs):
if "skill" in para.text.strip().lower():
j = i + 1

```
        while j < len(doc.paragraphs):
            text = doc.paragraphs[j].text.strip().lower()
            if "experience" in text:
                break

            doc.paragraphs[j].text = ""
            j += 1

        for idx, skill in enumerate(new_skills):
            if i + 1 + idx < len(doc.paragraphs):
                doc.paragraphs[i + 1 + idx].text = f"• {skill}"

        break
```

def update_experience(doc, tailored_experience):
for i, para in enumerate(doc.paragraphs):
job_title = para.text.strip()

```
    if job_title in tailored_experience:
        new_bullets = tailored_experience[job_title]

        employer_index = i + 1
        j = employer_index + 1

        while j < len(doc.paragraphs):
            text = doc.paragraphs[j].text.strip()

            if text in tailored_experience.keys() or text.lower() == "education":
                break

            j += 1

        for k in range(employer_index + 1, j):
            doc.paragraphs[k].text = ""

        for idx, bullet in enumerate(new_bullets):
            target_index = employer_index + 1 + idx

            if target_index < len(doc.paragraphs):
                doc.paragraphs[target_index].text = f"• {bullet}"
```

def process_application(job_description):
cv_text = read_docx(CV_PATH)
cover_text = read_docx(COVER_PATH)

```
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
tailored_cover = tailor_cover_letter_deep(
    cover_text,
    cv_text,
    job_description
)

cover_doc = Document(COVER_PATH)

if cover_doc.paragraphs:
    cover_doc.paragraphs[0].text = tailored_cover

cover_output = BytesIO()
cover_doc.save(cover_output)
cover_output.seek(0)

return cv_output, cover_output
```

@app.route("/")
def index():
return render_template("index.html")

@app.route("/tailor", methods=["POST"])
def tailor():
job_description = request.form.get("job_description", "").strip()

```
if not job_description:
    return jsonify({"error": "Please paste a job description"}), 400

cv_output, cover_output = process_application(job_description)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

return jsonify({
    "success": True,
    "cv_filename": f"Tailored_CV_{timestamp}.docx",
    "cover_filename": f"Tailored_Cover_{timestamp}.docx",
    "cv_data": base64.b64encode(cv_output.getvalue()).decode("utf-8"),
    "cover_data": base64.b64encode(cover_output.getvalue()).decode("utf-8")
})
```

@app.route("/health")
def health():
return jsonify({"status": "healthy"})

if **name** == "**main**":
port = int(os.environ.get("PORT", 8080))
app.run(host="0.0.0.0", port=port)
