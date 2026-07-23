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
            temperature=0.3,
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"DeepSeek API error: {str(e)}")
        raise e

def read_docx(file_path):
    """Extract text from .docx file preserving structure"""
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

def extract_sections_with_indices(text):
    """Extract sections with their starting indices"""
    lines = text.split('\n')
    sections = {
        "summary": {"start": -1, "end": -1, "content": ""},
        "skills": {"start": -1, "end": -1, "content": ""},
        "experience": {"start": -1, "end": -1, "content": ""},
        "education": {"start": -1, "end": -1, "content": ""}
    }
    
    section_keywords = {
        "summary": ["summary", "profile", "about"],
        "skills": ["skill", "core competencies", "expertise", "skill highlights"],
        "experience": ["experience", "employment", "work history"],
        "education": ["education", "academic", "qualification"]
    }
    
    current_section = None
    for i, line in enumerate(lines):
        line_lower = line.lower().strip()
        
        # Check if this line is a section header
        for section, keywords in section_keywords.items():
            if any(keyword in line_lower for keyword in keywords):
                if current_section:
                    sections[current_section]["end"] = i - 1
                current_section = section
                if sections[current_section]["start"] == -1:
                    sections[current_section]["start"] = i
                break
        
        # Add content to current section
        if current_section and sections[current_section]["start"] != -1:
            # Skip section header lines
            if sections[current_section]["start"] != i:
                clean_line = re.sub(r'\{#.*?\}', '', line)
                clean_line = re.sub(r'\.Styl\d+', '', clean_line)
                if clean_line.strip():
                    sections[current_section]["content"] += clean_line + "\n"
    
    return sections

def tailor_with_ai(cv_text, job_description):
    """Get tailored content from DeepSeek"""
    
    sections = extract_sections_with_indices(cv_text)
    
    prompt = f"""
You are a professional CV tailor. Rewrite ONLY the SUMMARY and SKILLS sections.

Current SUMMARY:
{sections['summary']['content'][:500]}

Current SKILLS:
{sections['skills']['content'][:500]}

JOB DESCRIPTION:
{job_description}

IMPORTANT RULES:
1. Rewrite ONLY the summary section to match the job
2. Rewrite ONLY the skills section to match the job
3. Keep ALL experience, education, and other sections EXACTLY as they are
4. Use keywords from the job description
5. NEVER add experience not in the CV
6. Keep the same professional tone

Return JSON:
{{
    "tailored_summary": "new summary (1-2 paragraphs)",
    "tailored_skills": "new skills list (bullet points or list)"
}}
"""
    
    return call_deepseek(prompt)

def tailor_cover_with_ai(cover_text, cv_text, job_description):
    """Generate tailored cover letter"""
    
    sections = extract_sections_with_indices(cv_text)
    
    prompt = f"""
You are a professional cover letter writer.

CV SUMMARY (use for context):
{sections['summary']['content'][:500]}

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE:
{cover_text}

INSTRUCTIONS:
1. Use the template structure
2. Highlight relevant experience from the CV
3. Match keywords from the job description
4. Include specific examples from the CV
5. Keep it professional and concise (3-4 paragraphs)
6. DO NOT fabricate experience

Return ONLY the cover letter text.
"""
    
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert cover letter writer. Return only the cover letter text."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Cover letter error: {str(e)}")
        return cover_text

def update_docx_sections(template_path, new_summary, new_skills):
    """
    Update ONLY summary and skills sections while preserving everything else
    """
    doc = Document(template_path)
    
    # Find section headers and their positions
    section_positions = []
    section_keywords = {
        "summary": ["summary", "profile", "about"],
        "skills": ["skill", "core competencies", "expertise", "skill highlights"],
        "experience": ["experience", "employment", "work history"],
        "education": ["education", "academic"]
    }
    
    for i, paragraph in enumerate(doc.paragraphs):
        text_lower = paragraph.text.lower().strip()
        for section, keywords in section_keywords.items():
            if any(keyword in text_lower for keyword in keywords):
                section_positions.append((section, i))
                break
    
    # Update summary section
    if new_summary and section_positions:
        summary_start = -1
        summary_end = -1
        skills_start = -1
        
        # Find summary and skills positions
        for section, idx in section_positions:
            if section == "summary":
                summary_start = idx
            elif section == "skills":
                skills_start = idx
        
        # If summary found, update it
        if summary_start != -1:
            # Find where summary ends (before skills or experience)
            summary_end = len(doc.paragraphs)
            for section, idx in section_positions:
                if idx > summary_start and section in ["skills", "experience", "education"]:
                    summary_end = idx
                    break
            
            # Replace summary content
            summary_lines = [p.strip() for p in new_summary.split('\n') if p.strip()]
            line_idx = summary_start + 1  # Start after header
            
            for i, line in enumerate(summary_lines):
                if line_idx + i < summary_end and line_idx + i < len(doc.paragraphs):
                    para = doc.paragraphs[line_idx + i]
                    if para.runs:
                        para.runs[0].text = line
                        for run in para.runs[1:]:
                            run.text = ""
                    else:
                        para.text = line
        
        # Update skills section
        if new_skills and skills_start != -1:
            # Find where skills ends
            skills_end = len(doc.paragraphs)
            for section, idx in section_positions:
                if idx > skills_start and section in ["experience", "education"]:
                    skills_end = idx
                    break
            
            # Replace skills content
            skills_lines = [p.strip() for p in new_skills.split('\n') if p.strip()]
            line_idx = skills_start + 1  # Start after header
            
            for i, line in enumerate(skills_lines):
                if line_idx + i < skills_end and line_idx + i < len(doc.paragraphs):
                    para = doc.paragraphs[line_idx + i]
                    if para.runs:
                        para.runs[0].text = line
                        for run in para.runs[1:]:
                            run.text = ""
                    else:
                        para.text = line
    
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
        print("🤖 Tailoring CV...")
        result = tailor_with_ai(cv_text, job_description)
        print(f"✅ CV tailored: {result.keys() if result else 'None'}")
    except Exception as e:
        return None, None, f"CV tailoring failed: {str(e)}"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_with_ai(cover_text, cv_text, job_description)
        print(f"✅ Cover letter generated")
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter generation failed: {str(e)}")
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = update_docx_sections(
            CV_PATH,
            result.get('tailored_summary', ''),
            result.get('tailored_skills', '')
        )
        print("✅ CV generated")
        
        # Generate cover letter
        cover_doc = Document(COVER_PATH)
        if tailored_cover:
            new_paragraphs = [p for p in tailored_cover.split('\n') if p.strip()]
            for i, paragraph in enumerate(cover_doc.paragraphs):
                if i < len(new_paragraphs):
                    if paragraph.runs:
                        paragraph.runs[0].text = new_paragraphs[i]
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
