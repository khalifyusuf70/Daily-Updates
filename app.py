import os
import json
import base64
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template, request, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
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
    """Call DeepSeek API with proper format"""
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
    """Extract text from .docx file"""
    try:
        doc = Document(file_path)
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text)
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return ""

def tailor_cv_deep(cv_text, job_description):
    """
    Deep tailoring of CV to match job description
    Rewrites Summary, Skills, AND Experience bullets
    """
    
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
3. Rewrite EACH experience bullet point to highlight the most relevant achievements.
4. Do NOT invent new jobs, achievements, numbers, or responsibilities.
5. Keep the same job titles, employers, and dates exactly as they are.
6. For each job, rewrite the bullet points to use keywords from the job description.
7. Return ONLY valid JSON.

Return JSON in this exact structure:
{{
    "tailored_summary": "new summary text here",
    "tailored_skills": "Skill 1\\nSkill 2\\nSkill 3\\n...",
    "tailored_experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "Rewritten bullet point 1",
            "Rewritten bullet point 2",
            "Rewritten bullet point 3"
        ],
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": [
            "Rewritten bullet point 1",
            "Rewritten bullet point 2"
        ]
    }}
}}
"""
    
    return call_deepseek(prompt)

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
    """Generate deeply tailored cover letter"""
    
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
    """Update summary section"""
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
                para = doc.paragraphs[i]
                para.clear()
        
        if summary_pos + 1 < len(doc.paragraphs):
            para = doc.paragraphs[summary_pos + 1]
            para.text = new_summary
        print("✅ Updated summary")

def update_skills(doc, new_skills):
    """Update skills section"""
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
                para = doc.paragraphs[i]
                para.clear()
        
        skills_lines = [s.strip() for s in new_skills.split('\n') if s.strip()]
        for i, skill in enumerate(skills_lines):
            if skills_pos + 1 + i < len(doc.paragraphs):
                para = doc.paragraphs[skills_pos + 1 + i]
                para.text = f"• {skill}"
        print(f"✅ Updated skills with {len(skills_lines)} lines")

def update_experience(doc, tailored_experience):
    """Update experience bullet points for each job"""
    
    print("\n📝 Updating experience bullets...")
    
    # Find all job titles in the document
    job_positions = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        # Match job title format (bold text with dates)
        if any(keyword in text.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager', 'director']):
            if '(' in text and ')' in text and ('20' in text or '19' in text):
                job_positions.append((i, text))
        # Also match the exact job titles from the CV
        if text in tailored_experience:
            if (i, text) not in job_positions:
                job_positions.append((i, text))
    
    print(f"📍 Found {len(job_positions)} job positions")
    
    for pos, job_title in job_positions:
        # Find matching job in tailored_experience
        matched_job = None
        for cv_job in tailored_experience.keys():
            # Check if the job title matches (fuzzy match)
            if job_title in cv_job or cv_job in job_title:
                matched_job = cv_job
                break
        
        if matched_job and tailored_experience[matched_job]:
            print(f"  ✅ Updating: {matched_job[:50]}...")
            new_bullets = tailored_experience[matched_job]
            
            # Find the end of this job section
            end_pos = len(doc.paragraphs)
            for next_pos, next_title in job_positions:
                if next_pos > pos:
                    end_pos = next_pos
                    break
            
            # Clear existing bullets
            for i in range(pos + 1, end_pos):
                if i < len(doc.paragraphs):
                    para = doc.paragraphs[i]
                    if para.text.strip() and not para.text.strip().startswith('•'):
                        # Keep the employer line
                        if 'Jubaland' in para.text or 'Ministry' in para.text or 'KIMS' in para.text:
                            continue
                    para.clear()
            
            # Insert new bullets
            insert_pos = pos + 1
            # Skip the employer line if it exists
            if insert_pos < len(doc.paragraphs):
                employer_text = doc.paragraphs[insert_pos].text.strip()
                if 'Jubaland' in employer_text or 'Ministry' in employer_text or 'KIMS' in employer_text:
                    insert_pos += 1
            
            for i, bullet in enumerate(new_bullets):
                if insert_pos + i < len(doc.paragraphs):
                    para = doc.paragraphs[insert_pos + i]
                    para.text = f"• {bullet}"
            
            print(f"     Inserted {len(new_bullets)} bullets")

def create_tailored_cv(template_path, new_summary, new_skills, tailored_experience):
    """Create tailored CV with all sections updated"""
    
    doc = Document(template_path)
    
    update_summary(doc, new_summary)
    update_skills(doc, new_skills)
    update_experience(doc, tailored_experience)
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def process_application(job_description):
    """Process the entire application tailoring"""
    
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
        print(f"📝 New Skills: {tailored_skills[:100]}...")
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
        
        # Generate cover letter
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
