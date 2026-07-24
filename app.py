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

def call_deepseek(prompt):
    try:
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
    except Exception as e:
        print(f"DeepSeek API error: {str(e)}")
        raise e

def read_docx(file_path):
    try:
        doc = Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return ""

def tailor_cv_deep(cv_text, job_description):
    """
    Rewrite Summary, Skills, and Experience bullets
    """
    prompt = f"""
You are a professional CV tailoring expert.

MASTER CV:
{cv_text}

JOB DESCRIPTION:
{job_description}

TASK:
Rewrite the CV to align with the job description, while being 100% truthful.

RULES:
1. Rewrite the professional summary.
2. Rewrite the skills section as a list.
3. For EACH job, rewrite the bullet points to highlight relevant achievements.
4. Do NOT invent new jobs, achievements, or numbers.
5. Keep job titles, employers, and dates exactly as they appear in the Master CV.
6. Use the EXACT job title strings from the Master CV as keys for the experience object.
7. Return ONLY valid JSON.

Return JSON:
{{
    "tailored_summary": "new summary",
    "tailored_skills": "skill 1\\nskill 2\\n...",
    "tailored_experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "Rewritten bullet 1",
            "Rewritten bullet 2"
        ],
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": [
            "Rewritten bullet"
        ],
        "Chief Operations Officer | Jul 2016 -- Jul 2021": [
            "Rewritten bullet"
        ]
    }}
}}
"""
    return call_deepseek(prompt)

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
    prompt = f"""
You are a professional cover letter writer.

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

Create a 3-paragraph cover letter that perfectly matches this job.
Return ONLY the cover letter text.
"""
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert cover letter writer. Return only the cover letter text."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.4
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Cover letter error: {str(e)}")
        return cover_text

def update_summary(doc, new_summary):
    summary_pos = -1
    skills_pos = -1
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'summary' in text and len(text) < 30:
            summary_pos = i
        elif 'skill' in text and len(text) < 30:
            skills_pos = i
            break
    if summary_pos != -1 and new_summary:
        end_pos = skills_pos if skills_pos > summary_pos else len(doc.paragraphs)
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].clear()
        if summary_pos + 1 < len(doc.paragraphs):
            doc.paragraphs[summary_pos + 1].text = new_summary
        print("✅ Updated summary")

def update_skills(doc, new_skills):
    skills_pos = -1
    experience_pos = -1
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'skill' in text and len(text) < 30:
            skills_pos = i
        elif 'experience' in text and len(text) < 30:
            experience_pos = i
            break
    if skills_pos != -1 and new_skills:
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].clear()
        skills_lines = [s.strip() for s in new_skills.split('\n') if s.strip()]
        for i, skill in enumerate(skills_lines):
            if skills_pos + 1 + i < len(doc.paragraphs):
                doc.paragraphs[skills_pos + 1 + i].text = f"• {skill}"
        print(f"✅ Updated skills with {len(skills_lines)} lines")

def update_experience(doc, tailored_experience):
    """
    Update only the bullet points under each job.
    Preserves job title and employer lines.
    """
    print("\n📝 Updating experience bullets...")

    # First, collect all job title indices and the exact job title text
    job_indices = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        # Match job title pattern: contains parentheses and year ranges
        if '(' in text and ')' in text and any(str(y) in text for y in range(2010, 2030)):
            job_indices.append((i, text))
        # Also match the exact job titles we expect
        for job_title in tailored_experience.keys():
            if job_title in text:
                job_indices.append((i, text))
                break

    # Remove duplicates
    seen = set()
    unique_jobs = []
    for idx, title in job_indices:
        if idx not in seen:
            seen.add(idx)
            unique_jobs.append((idx, title))

    print(f"📍 Found {len(unique_jobs)} job sections")

    # Process each job
    for pos, job_title in unique_jobs:
        # Find matching key in tailored_experience
        matched_key = None
        for key in tailored_experience.keys():
            if key in job_title or job_title in key:
                matched_key = key
                break
        if not matched_key:
            print(f"⚠️ No match for job: {job_title[:50]}")
            continue

        new_bullets = tailored_experience[matched_key]
        if not new_bullets:
            continue

        # Find the end of this job section (next job or Education)
        end_pos = len(doc.paragraphs)
        for next_pos, _ in unique_jobs:
            if next_pos > pos:
                end_pos = next_pos
                break
        # Also stop at Education section
        for i in range(pos + 1, end_pos):
            if 'education' in doc.paragraphs[i].text.lower():
                end_pos = i
                break

        print(f"  ✅ Updating: {matched_key[:50]}...")

        # We need to preserve the employer line (the line right after job title if it has company name)
        # and any lines that are not bullet points.

        # Collect all paragraph indices that are bullet points within this section
        bullet_indices = []
        for i in range(pos + 1, end_pos):
            para = doc.paragraphs[i]
            text = para.text.strip()
            if text.startswith('-') or text.startswith('•') or text.startswith('*'):
                bullet_indices.append(i)
            # If it's a non-empty line that is not a bullet and not the employer, we keep it (like employer line)

        # Replace bullet points in order
        # We'll clear each bullet and set new text
        for idx, bullet_text in zip(bullet_indices, new_bullets):
            if idx < len(doc.paragraphs):
                para = doc.paragraphs[idx]
                # Preserve the bullet symbol style (if any)
                if para.runs:
                    para.runs[0].text = f"• {bullet_text}"
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = f"• {bullet_text}"

        # If there are more new bullets than old bullet slots, add them at the end
        if len(new_bullets) > len(bullet_indices):
            insert_pos = bullet_indices[-1] + 1 if bullet_indices else end_pos
            for extra in new_bullets[len(bullet_indices):]:
                if insert_pos < len(doc.paragraphs):
                    para = doc.paragraphs[insert_pos]
                    if para.runs:
                        para.runs[0].text = f"• {extra}"
                        for run in para.runs[1:]:
                            run.text = ""
                    else:
                        para.text = f"• {extra}"
                    insert_pos += 1

        print(f"     Inserted {len(new_bullets)} bullets")

def create_tailored_cv(template_path, new_summary, new_skills, tailored_experience):
    doc = Document(template_path)
    update_summary(doc, new_summary)
    update_skills(doc, new_skills)
    update_experience(doc, tailored_experience)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def process_application(job_description):
    if not os.path.exists(CV_PATH):
        return None, None, f"CV file not found: {CV_PATH}"
    if not os.path.exists(COVER_PATH):
        return None, None, f"Cover letter template not found: {COVER_PATH}"

    cv_text = read_docx(CV_PATH)
    cover_text = read_docx(COVER_PATH)
    if not cv_text or not cover_text:
        return None, None, "Failed to read documents"

    try:
        print("🤖 Tailoring CV with DeepSeek...")
        result = tailor_cv_deep(cv_text, job_description)
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        tailored_experience = result.get('tailored_experience', {})
        print(f"📝 New Summary: {tailored_summary[:100]}...")
        print(f"📝 Jobs to update: {list(tailored_experience.keys())}")
    except Exception as e:
        return None, None, f"AI tailoring failed: {str(e)}"

    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter_deep(cover_text, cv_text, job_description)
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter failed: {str(e)}")

    try:
        print("📄 Generating tailored CV...")
        cv_output = create_tailored_cv(CV_PATH, tailored_summary, tailored_skills, tailored_experience)

        # Cover letter
        cover_doc = Document(COVER_PATH)
        if tailored_cover:
            new_paragraphs = [p for p in tailored_cover.split('\n') if p.strip()]
            for i, paragraph in enumerate(cover_doc.paragraphs):
                if i < len(new_paragraphs):
                    if paragraph.runs:
                        paragraph.runs[0].text = new_paragraphs[i]
                        for run in paragraph.runs[1:]:
                            run.text = ""
                    else:
                        paragraph.text = new_paragraphs[i]
        cover_output = BytesIO()
        cover_doc.save(cover_output)
        cover_output.seek(0)
        print("✅ Documents generated")
        return cv_output, cover_output, "Success"
    except Exception as e:
        return None, None, f"Document generation error: {str(e)}"

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        return f"Error loading template: {str(e)}", 500

@app.route('/tailor', methods=['POST'])
def tailor():
    try:
        job_description = request.form.get('job_description', '').strip()
        if not job_description:
            return jsonify({'error': 'Please paste a job description'}), 400
        print(f"📝 Processing job: {len(job_description)} characters")
        cv_output, cover_output, message = process_application(job_description)
        if cv_output and cover_output:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            cv_base64 = base64.b64encode(cv_output.getvalue()).decode('utf-8')
            cover_base64 = base64.b64encode(cover_output.getvalue()).decode('utf-8')
            return jsonify({
                'success': True,
                'message': 'Documents tailored successfully!',
                'cv_filename': f'Tailored_CV_{timestamp}.docx',
                'cover_filename': f'Tailored_Cover_{timestamp}.docx',
                'cv_data': cv_base64,
                'cover_data': cover_base64
            })
        else:
            return jsonify({'error': message}), 500
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    print(f"🚀 Starting CV Tailor on port {port}")
    print(f"📂 Files: {os.listdir('.')}")
    app.run(host='0.0.0.0', port=port, debug=False)
