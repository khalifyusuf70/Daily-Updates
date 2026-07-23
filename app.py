import os
import json
import base64
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template, request, jsonify
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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

def get_full_docx_content(file_path):
    """Get full DOCX content with paragraphs"""
    doc = Document(file_path)
    content = []
    for para in doc.paragraphs:
        content.append(para.text)
    return content, doc

def tailor_summary_and_skills(cv_text, job_description):
    """Get tailored summary and skills from AI"""
    
    # Extract current summary
    lines = cv_text.split('\n')
    current_summary = ""
    in_summary = False
    
    for line in lines:
        line_lower = line.lower().strip()
        if 'summary' in line_lower and len(line) < 50:
            in_summary = True
            continue
        elif 'skill' in line_lower and len(line) < 50:
            in_summary = False
            continue
        elif 'experience' in line_lower and len(line) < 50:
            in_summary = False
            continue
        
        if in_summary and line.strip():
            clean = re.sub(r'\{#.*?\}', '', line)
            clean = re.sub(r'\.Styl\d+', '', clean)
            if clean.strip() and not clean.startswith('#'):
                current_summary += clean.strip() + " "
    
    if not current_summary:
        current_summary = "Senior professional with 10+ years of experience in government and international development."
    
    prompt = f"""
You are a professional CV tailor. Rewrite ONLY the SUMMARY section.

CURRENT SUMMARY:
{current_summary}

JOB DESCRIPTION:
{job_description}

INSTRUCTIONS:
Write a 4-6 sentence summary that PERFECTLY matches this job. Use keywords from the job description.

Also, list 10-12 key skills for this role in a comma-separated list.

Return ONLY JSON:
{{
    "tailored_summary": "new summary here",
    "tailored_skills": "skill 1, skill 2, skill 3, skill 4, skill 5, skill 6, skill 7, skill 8, skill 9, skill 10"
}}
"""
    
    return call_deepseek(prompt)

def create_tailored_cv(original_path, new_summary, new_skills, output_path=None):
    """
    Create a new CV by copying the original and updating only the text
    """
    # Load original document
    doc = Document(original_path)
    
    # Find and update the summary section
    summary_pos = -1
    skills_pos = -1
    experience_pos = -1
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'summary' in text and len(text) < 30:
            summary_pos = i
        elif 'skill' in text and len(text) < 30:
            skills_pos = i
        elif 'experience' in text and len(text) < 30:
            experience_pos = i
            break
    
    print(f"📍 Found: Summary={summary_pos}, Skills={skills_pos}, Experience={experience_pos}")
    
    # Update summary
    if summary_pos != -1 and new_summary:
        # Find the next section header
        end_pos = skills_pos if skills_pos > summary_pos else experience_pos if experience_pos > summary_pos else len(doc.paragraphs)
        
        # Clear existing summary text (keep the header)
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                if para.text.strip():
                    para.clear()
        
        # Insert new summary as a single paragraph
        if summary_pos + 1 < len(doc.paragraphs):
            para = doc.paragraphs[summary_pos + 1]
            para.text = new_summary
    
    # Update skills - Remove the old skills table and add new skills as text
    if skills_pos != -1 and new_skills:
        # Find end of skills section
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        
        # Clear everything from skills header to experience header
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                para = doc.paragraphs[i]
                para.clear()
        
        # Insert new skills as a bulleted list
        skills_list = [s.strip() for s in new_skills.split(',') if s.strip()]
        
        for i, skill in enumerate(skills_list):
            if skills_pos + 1 + i < len(doc.paragraphs):
                para = doc.paragraphs[skills_pos + 1 + i]
                para.text = f"• {skill}"
    
    # Save to BytesIO
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
        print("🤖 Getting tailored summary and skills...")
        result = tailor_summary_and_skills(cv_text, job_description)
        
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        
        print(f"📝 New Summary: {tailored_summary[:100]}...")
        print(f"📝 New Skills: {tailored_skills[:100]}...")
        
    except Exception as e:
        return None, None, f"AI tailoring failed: {str(e)}"
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = create_tailored_cv(CV_PATH, tailored_summary, tailored_skills)
        print("✅ CV generated")
        
        # Generate cover letter - use original template
        cover_doc = Document(COVER_PATH)
        cover_output = BytesIO()
        cover_doc.save(cover_output)
        cover_output.seek(0)
        print("✅ Cover letter generated")
        
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
