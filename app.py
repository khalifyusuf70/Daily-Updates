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
    """Deep tailoring of CV to match job description"""
    
    # Extract current summary and skills - SIMPLER APPROACH
    lines = cv_text.split('\n')
    current_summary = ""
    current_skills = ""
    in_summary = False
    in_skills = False
    
    for i, line in enumerate(lines):
        line_lower = line.lower().strip()
        
        # Detect sections
        if 'summary' in line_lower and '#' in line:
            in_summary = True
            in_skills = False
            continue
        elif 'skill' in line_lower and '#' in line:
            in_summary = False
            in_skills = True
            continue
        elif 'experience' in line_lower and '#' in line:
            in_summary = False
            in_skills = False
            continue
        elif 'education' in line_lower and '#' in line:
            in_summary = False
            in_skills = False
            continue
        
        # Collect content
        if in_summary and line.strip() and not line.startswith('#'):
            clean = re.sub(r'\{#.*?\}', '', line)
            clean = re.sub(r'\.Styl\d+', '', clean)
            if clean.strip():
                current_summary += clean.strip() + " "
        elif in_skills and line.strip() and not line.startswith('#'):
            clean = re.sub(r'\{#.*?\}', '', line)
            clean = re.sub(r'\.Styl\d+', '', clean)
            if clean.strip() and not clean.startswith('+'):
                current_skills += clean.strip() + "\n"
    
    print(f"📝 Extracted Summary: {current_summary[:100]}...")
    print(f"📝 Extracted Skills: {current_skills[:100]}...")
    
    if not current_summary:
        current_summary = "Senior professional with 10+ years of experience in government and international development."
    if not current_skills:
        current_skills = "Resource Mobilization, Grants Management, Stakeholder Engagement, Strategic Leadership"
    
    prompt = f"""
You are a professional CV tailor. Rewrite ONLY the SUMMARY and SKILLS sections.

CURRENT SUMMARY (1 paragraph):
{current_summary}

CURRENT SKILLS (list):
{current_skills}

JOB DESCRIPTION:
{job_description}

INSTRUCTIONS:
1. SUMMARY: Write 1 paragraph (4-6 sentences) that PERFECTLY matches this job.
2. SKILLS: List 10-12 key skills as a bulleted list matching this job.

Return ONLY JSON:
{{
    "tailored_summary": "4-6 sentence paragraph matching the job",
    "tailored_skills": "bullet 1\\nbullet 2\\nbullet 3\\n..."
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

def update_docx_sections(template_path, new_summary, new_skills):
    """
    SAFELY update only the summary and skills sections.
    NEVER touch experience or education sections.
    """
    doc = Document(template_path)
    
    # Find the exact positions of section headers
    summary_pos = -1
    skills_pos = -1
    experience_pos = -1
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'summary' in text and '#' in para.text:
            summary_pos = i
        elif 'skill' in text and '#' in para.text:
            skills_pos = i
        elif 'experience' in text and '#' in para.text:
            experience_pos = i
            break  # Stop once we find experience
    
    print(f"📍 Summary at: {summary_pos}, Skills at: {skills_pos}, Experience at: {experience_pos}")
    
    # --- UPDATE SUMMARY ---
    if summary_pos != -1 and new_summary:
        # Find where summary ends (before skills or experience)
        end_pos = len(doc.paragraphs)
        if skills_pos != -1 and skills_pos > summary_pos:
            end_pos = skills_pos
        elif experience_pos != -1 and experience_pos > summary_pos:
            end_pos = experience_pos
        
        # Clear all content between summary header and next header
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                if para.runs:
                    para.runs[0].text = ""
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = ""
        
        # Insert new summary as a single paragraph
        if summary_pos + 1 < len(doc.paragraphs):
            para = doc.paragraphs[summary_pos + 1]
            if para.runs:
                para.runs[0].text = new_summary
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.text = new_summary
        print(f"✅ Updated summary")
    
    # --- UPDATE SKILLS ---
    if skills_pos != -1 and new_skills:
        # Find where skills ends (before experience)
        end_pos = len(doc.paragraphs)
        if experience_pos != -1 and experience_pos > skills_pos:
            end_pos = experience_pos
        
        # Clear all content between skills header and experience header
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                if para.runs:
                    para.runs[0].text = ""
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = ""
        
        # Insert new skills
        skills_lines = [p.strip() for p in new_skills.split('\n') if p.strip()]
        
        for i, line in enumerate(skills_lines):
            if skills_pos + 1 + i < len(doc.paragraphs):
                para = doc.paragraphs[skills_pos + 1 + i]
                if para.runs:
                    para.runs[0].text = line
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = line
        print(f"✅ Updated skills with {len(skills_lines)} lines")
    
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
        print("🤖 Tailoring CV...")
        result = tailor_cv_deep(cv_text, job_description)
        
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        
        print(f"📝 New Summary: {tailored_summary[:100]}...")
        print(f"📝 New Skills: {tailored_skills[:100]}...")
        
    except Exception as e:
        return None, None, f"CV tailoring failed: {str(e)}"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter_deep(cover_text, cv_text, job_description)
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter failed: {str(e)}")
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = update_docx_sections(CV_PATH, tailored_summary, tailored_skills)
        
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
