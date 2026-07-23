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

def extract_summary_and_skills(text):
    """Extract summary and skills sections from CV text"""
    lines = text.split('\n')
    summary = []
    skills = []
    current_section = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        line_lower = line.lower()
        
        # Detect section headers
        if 'summary' in line_lower or 'profile' in line_lower:
            current_section = 'summary'
            continue
        elif 'skill' in line_lower or 'core competency' in line_lower or 'expertise' in line_lower:
            current_section = 'skills'
            continue
        elif 'experience' in line_lower or 'education' in line_lower:
            current_section = None
            continue
        
        # Add content to current section
        if current_section == 'summary':
            # Clean up formatting artifacts
            clean_line = re.sub(r'\{#.*?\}', '', line)
            clean_line = re.sub(r'\.Styl\d+', '', clean_line)
            if clean_line.strip():
                summary.append(clean_line)
        elif current_section == 'skills':
            # Clean up formatting artifacts
            clean_line = re.sub(r'\{#.*?\}', '', line)
            clean_line = re.sub(r'\.Styl\d+', '', clean_line)
            if clean_line.strip():
                skills.append(clean_line)
    
    return {
        'summary': '\n'.join(summary) if summary else '',
        'skills': '\n'.join(skills) if skills else ''
    }

def tailor_cv_deep(cv_text, job_description):
    """Deep tailoring of CV to match job description"""
    
    sections = extract_summary_and_skills(cv_text)
    
    print(f"📝 Original Summary: {sections['summary'][:200]}...")
    print(f"📝 Original Skills: {sections['skills'][:200]}...")
    
    prompt = f"""
You are a professional CV tailor. Your task is to rewrite the SUMMARY and SKILLS sections to PERFECTLY match the job description.

CURRENT SUMMARY (rewrite this):
{sections['summary']}

CURRENT SKILLS (rewrite this):
{sections['skills']}

JOB DESCRIPTION:
{job_description}

CRITICAL INSTRUCTIONS:

For the SUMMARY:
1. Must mention: Government Affairs, Public Affairs, Policy Advocacy
2. Must mention: Stakeholder Engagement, Partnership Development
3. Must mention: Sub-Saharan Africa experience
4. Must mention: Business Development or Market Access
5. Must mention: Public Relations or Strategic Communications
6. Use keywords from the job description naturally
7. Keep professional tone, 2-3 sentences

For the SKILLS:
1. MUST include: Government Affairs & Policy Advocacy
2. MUST include: Stakeholder Engagement & Partnership Development
3. MUST include: Public Relations & Strategic Communications
4. MUST include: Market Access & Business Development
5. MUST include: Regulatory Analysis & Policy Monitoring
6. MUST include: Donor Relations & Resource Mobilization
7. MUST include: Cross-functional Collaboration & Advisory
8. Format as a bulleted list or comma-separated

Return JSON:
{{
    "tailored_summary": "new summary here (2-3 sentences)",
    "tailored_skills": "new skills list here"
}}
"""
    
    return call_deepseek(prompt)

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
    """Generate deeply tailored cover letter"""
    
    sections = extract_summary_and_skills(cv_text)
    
    prompt = f"""
You are a professional cover letter writer. Create a compelling cover letter for this job.

CV SUMMARY (for context):
{sections['summary']}

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

INSTRUCTIONS:
1. Write a 3-4 paragraph cover letter
2. Open with enthusiasm for Novonesis and biosolutions
3. Highlight your experience in government affairs and public policy
4. Mention specific achievements from your CV that match the job
5. Use keywords: biosolutions, policy advocacy, government relations, market access
6. Address the three areas: Business Enablement, Government Affairs, Public Relations
7. Keep it professional and compelling
8. DO NOT fabricate any experience

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
    """Update ONLY summary and skills sections in the DOCX"""
    doc = Document(template_path)
    
    print(f"📝 New Summary: {new_summary[:100]}...")
    print(f"📝 New Skills: {new_skills[:100]}...")
    
    # Find section headers and update content
    in_summary = False
    in_skills = False
    summary_updated = False
    skills_updated = False
    
    for paragraph in doc.paragraphs:
        text_lower = paragraph.text.lower().strip()
        
        # Detect section headers
        if 'summary' in text_lower or 'profile' in text_lower:
            in_summary = True
            in_skills = False
            continue
        elif 'skill' in text_lower or 'core competency' in text_lower or 'expertise' in text_lower:
            in_summary = False
            in_skills = True
            continue
        elif 'experience' in text_lower or 'education' in text_lower:
            in_summary = False
            in_skills = False
            continue
        
        # Update summary section
        if in_summary and not summary_updated and new_summary:
            if paragraph.runs:
                paragraph.runs[0].text = new_summary
                for run in paragraph.runs[1:]:
                    run.text = ""
                summary_updated = True
            else:
                paragraph.text = new_summary
                summary_updated = True
        
        # Update skills section
        elif in_skills and not skills_updated and new_skills:
            if paragraph.runs:
                paragraph.runs[0].text = new_skills
                for run in paragraph.runs[1:]:
                    run.text = ""
                skills_updated = True
            else:
                paragraph.text = new_skills
                skills_updated = True
    
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
    
    print("📝 Extracting current sections...")
    current_sections = extract_summary_and_skills(cv_text)
    print(f"✅ Current Summary: {current_sections['summary'][:100]}...")
    print(f"✅ Current Skills: {current_sections['skills'][:100]}...")
    
    try:
        print("🤖 Tailoring CV with DeepSeek...")
        result = tailor_cv_deep(cv_text, job_description)
        print(f"✅ Result keys: {result.keys() if result else 'None'}")
        
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        
        print(f"📝 Tailored Summary: {tailored_summary[:100]}...")
        print(f"📝 Tailored Skills: {tailored_skills[:100]}...")
        
    except Exception as e:
        return None, None, f"CV tailoring failed: {str(e)}"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter_deep(cover_text, cv_text, job_description)
        print(f"✅ Cover letter: {len(tailored_cover)} characters")
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter failed: {str(e)}")
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = update_docx_sections(CV_PATH, tailored_summary, tailored_skills)
        print("✅ CV generated")
        
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
