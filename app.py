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

def get_docx_paragraphs(file_path):
    """Get all paragraphs from DOCX with their indices"""
    doc = Document(file_path)
    return [(i, para) for i, para in enumerate(doc.paragraphs)]

def extract_sections_by_header(doc):
    """Extract sections based on header detection"""
    sections = {
        'summary': {'start': -1, 'end': -1, 'content': ''},
        'skills': {'start': -1, 'end': -1, 'content': ''},
        'experience': {'start': -1, 'end': -1, 'content': ''},
        'education': {'start': -1, 'end': -1, 'content': ''}
    }
    
    current_section = None
    section_headers = {
        'summary': ['summary', 'profile', 'professional summary'],
        'skills': ['skill', 'core competencies', 'expertise', 'skill highlights'],
        'experience': ['experience', 'employment', 'work history', 'professional experience'],
        'education': ['education', 'academic', 'qualification']
    }
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip().lower()
        
        # Check if this is a section header
        found_section = None
        for section, keywords in section_headers.items():
            if any(keyword in text for keyword in keywords):
                found_section = section
                break
        
        if found_section:
            current_section = found_section
            if sections[current_section]['start'] == -1:
                sections[current_section]['start'] = i
            continue
        
        # If we're in a section, collect content
        if current_section and sections[current_section]['start'] != -1:
            # Skip the header line itself
            if sections[current_section]['start'] != i:
                clean_text = re.sub(r'\{#.*?\}', '', para.text)
                clean_text = re.sub(r'\.Styl\d+', '', clean_text)
                if clean_text.strip():
                    sections[current_section]['content'] += clean_text + "\n"
    
    # Set end positions
    section_names = list(sections.keys())
    for i, section in enumerate(section_names):
        if i < len(section_names) - 1:
            next_section = section_names[i + 1]
            if sections[next_section]['start'] != -1:
                sections[section]['end'] = sections[next_section]['start']
        if sections[section]['end'] == -1:
            sections[section]['end'] = len(doc.paragraphs)
    
    return sections

def tailor_cv_deep(cv_text, job_description):
    """Deep tailoring of CV to match job description"""
    
    # Extract current summary and skills
    current_summary = ""
    current_skills = ""
    
    lines = cv_text.split('\n')
    in_summary = False
    in_skills = False
    
    for line in lines:
        line_lower = line.lower().strip()
        if 'summary' in line_lower or 'profile' in line_lower:
            in_summary = True
            in_skills = False
            continue
        elif 'skill' in line_lower or 'core competency' in line_lower:
            in_summary = False
            in_skills = True
            continue
        elif 'experience' in line_lower or 'education' in line_lower:
            in_summary = False
            in_skills = False
            continue
        
        if in_summary and line.strip():
            clean = re.sub(r'\{#.*?\}', '', line)
            clean = re.sub(r'\.Styl\d+', '', clean)
            if clean.strip():
                current_summary += clean + "\n"
        elif in_skills and line.strip():
            clean = re.sub(r'\{#.*?\}', '', line)
            clean = re.sub(r'\.Styl\d+', '', clean)
            if clean.strip():
                current_skills += clean + "\n"
    
    print(f"📝 Extracted Summary: {current_summary[:100]}...")
    print(f"📝 Extracted Skills: {current_skills[:100]}...")
    
    # If we didn't find sections, use defaults
    if not current_summary:
        current_summary = "Senior professional with 10+ years of experience in government and international development sectors. Expert in strategic leadership, programme growth, resource mobilization, and donor engagement."
    if not current_skills:
        current_skills = "Resource Mobilization, Grants Management, Stakeholder Engagement, Strategic Leadership, Donor Relations, Partnership Development"
    
    # Analyze job description to determine role type
    prompt = f"""
You are a professional CV tailor. Rewrite the SUMMARY and SKILLS sections to PERFECTLY match this job.

CURRENT SUMMARY (REWRITE THIS):
{current_summary}

CURRENT SKILLS (REWRITE THIS):
{current_skills}

JOB DESCRIPTION:
{job_description}

The new summary and skills must be specifically tailored to this exact job.

Return ONLY JSON:
{{
    "tailored_summary": "new summary that directly matches the job requirements",
    "tailored_skills": "new skills list that directly matches the job requirements"
}}
"""
    
    return call_deepseek(prompt)

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
    """Generate deeply tailored cover letter"""
    
    prompt = f"""
You are a professional cover letter writer. Create a compelling cover letter for this job.

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

Return ONLY the cover letter text (no JSON).
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
    
    # First, find section header positions
    header_indices = {}
    section_headers = {
        'summary': ['summary', 'profile'],
        'skills': ['skill', 'core', 'expertise'],
        'experience': ['experience', 'employment'],
        'education': ['education', 'academic']
    }
    
    for i, para in enumerate(doc.paragraphs):
        text_lower = para.text.lower().strip()
        for section, keywords in section_headers.items():
            if any(keyword in text_lower for keyword in keywords):
                if section not in header_indices:
                    header_indices[section] = i
                break
    
    print(f"📝 Found headers: {header_indices}")
    
    # Update summary section
    if 'summary' in header_indices and new_summary:
        summary_idx = header_indices['summary']
        # Find next section header
        next_header = len(doc.paragraphs)
        for section, idx in header_indices.items():
            if idx > summary_idx and idx < next_header:
                next_header = idx
        
        # Replace summary content (skip the header line)
        summary_lines = [p.strip() for p in new_summary.split('\n') if p.strip()]
        line_idx = summary_idx + 1
        
        for i, line in enumerate(summary_lines):
            if line_idx + i < next_header and line_idx + i < len(doc.paragraphs):
                para = doc.paragraphs[line_idx + i]
                if para.runs:
                    para.runs[0].text = line
                    for run in para.runs[1:]:
                        run.text = ""
                else:
                    para.text = line
        print(f"✅ Updated summary with {len(summary_lines)} lines")
    
    # Update skills section
    if 'skills' in header_indices and new_skills:
        skills_idx = header_indices['skills']
        # Find next section header
        next_header = len(doc.paragraphs)
        for section, idx in header_indices.items():
            if idx > skills_idx and idx < next_header:
                next_header = idx
        
        # Replace skills content (skip the header line)
        skills_lines = [p.strip() for p in new_skills.split('\n') if p.strip()]
        line_idx = skills_idx + 1
        
        for i, line in enumerate(skills_lines):
            if line_idx + i < next_header and line_idx + i < len(doc.paragraphs):
                para = doc.paragraphs[line_idx + i]
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
        print("🤖 Tailoring CV with DeepSeek...")
        result = tailor_cv_deep(cv_text, job_description)
        print(f"✅ Result: {result.keys() if result else 'None'}")
        
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        
        print(f"📝 New Summary: {tailored_summary[:100]}...")
        print(f"📝 New Skills: {tailored_skills[:100]}...")
        
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
