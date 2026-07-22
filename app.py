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
    print("⚠️ DEEPSEEK_API_KEY not set. Please set it in environment variables.")

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
    """Extract text from .docx file with paragraph structure"""
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

def extract_sections(text):
    """Extract sections from CV text more intelligently"""
    sections = {
        "summary": "",
        "skills": "",
        "experience": "",
        "education": ""
    }
    
    lines = text.split('\n')
    current_section = None
    
    # Keywords for section detection
    section_keywords = {
        "summary": ["summary", "profile", "about"],
        "skills": ["skill", "core competencies", "expertise"],
        "experience": ["experience", "employment", "work history"],
        "education": ["education", "academic", "qualification"]
    }
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        line_lower = line.lower()
        section_found = False
        
        # Check if this line is a section header
        for section, keywords in section_keywords.items():
            if any(keyword in line_lower for keyword in keywords):
                current_section = section
                section_found = True
                break
        
        # If not a header, add to current section
        if not section_found and current_section:
            # Clean up formatting artifacts
            clean_line = re.sub(r'\{#.*?\}', '', line)  # Remove {#...}
            clean_line = re.sub(r'\.Styl\d+', '', clean_line)  # Remove .Styl classes
            if clean_line.strip():
                sections[current_section] += clean_line + "\n"
    
    return sections

def tailor_cv_with_ai(cv_text, job_description):
    """Get tailored content from DeepSeek with specific instructions"""
    
    sections = extract_sections(cv_text)
    
    prompt = f"""
You are a professional CV tailor. Your task is to rewrite ONLY the SUMMARY and SKILLS sections.

Current CV Sections:
SUMMARY:
{sections['summary']}

SKILLS:
{sections['skills']}

EXPERIENCE (DO NOT CHANGE):
{sections['experience'][:500]}...

JOB DESCRIPTION:
{job_description}

INSTRUCTIONS:
1. Rewrite ONLY the SUMMARY section to perfectly match the job requirements
2. Rewrite ONLY the SKILLS section to highlight skills from the job description
3. Keep ALL experience, education, and other sections EXACTLY as they are
4. Use keywords and phrases from the job description naturally
5. NEVER add experience or qualifications not in the CV
6. Keep the same tone and style as the original

Return JSON:
{{
    "tailored_summary": "new summary here (1-2 paragraphs)",
    "tailored_skills": "new skills here (bullet points or list)"
}}
"""
    
    return call_deepseek(prompt)

def tailor_cover_with_ai(cover_text, cv_text, job_description):
    """Generate tailored cover letter"""
    
    sections = extract_sections(cv_text)
    
    prompt = f"""
You are a professional cover letter writer. Create a tailored cover letter.

CV SUMMARY (use this for context):
{sections['summary']}

JOB DESCRIPTION:
{job_description}

COVER LETTER TEMPLATE (use this structure):
{cover_text}

INSTRUCTIONS:
1. Write a compelling cover letter using the template structure
2. Highlight the most relevant experience from the CV
3. Match keywords from the job description
4. Keep it professional and concise (3-4 paragraphs)
5. Include specific examples from the CV that match the job requirements
6. DO NOT fabricate any experience or qualifications

Return ONLY the cover letter text, no JSON.
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
        return cover_text  # Return original if fails

def preserve_formatting_and_tailor(template_path, new_summary, new_skills, output_path=None):
    """
    Preserve exact formatting while updating only summary and skills sections
    """
    doc = Document(template_path)
    
    summary_found = False
    skills_found = False
    
    # First pass: find section headers
    section_positions = []
    for i, paragraph in enumerate(doc.paragraphs):
        text_lower = paragraph.text.lower().strip()
        if any(keyword in text_lower for keyword in ['summary', 'profile']):
            section_positions.append(('summary', i))
        elif any(keyword in text_lower for keyword in ['skill', 'core competencies']):
            section_positions.append(('skills', i))
        elif any(keyword in text_lower for keyword in ['experience', 'education']):
            # Stop processing after experience section
            break
    
    # Process sections
    for section_type, start_idx in section_positions:
        # Find the end of this section (next section header or end of document)
        end_idx = len(doc.paragraphs)
        for next_type, next_idx in section_positions:
            if next_idx > start_idx:
                end_idx = next_idx
                break
        
        # Update the section content
        if section_type == 'summary' and new_summary:
            # Replace the summary paragraph(s)
            summary_lines = [p.strip() for p in new_summary.split('\n') if p.strip()]
            for i, line in enumerate(summary_lines):
                if start_idx + i < end_idx and start_idx + i < len(doc.paragraphs):
                    para = doc.paragraphs[start_idx + i]
                    if para.runs:
                        para.runs[0].text = line
                        # Clear other runs
                        for run in para.runs[1:]:
                            run.text = ""
                    else:
                        # If no runs, create one
                        para.text = line
        
        elif section_type == 'skills' and new_skills:
            # Replace the skills section
            skills_lines = [p.strip() for p in new_skills.split('\n') if p.strip()]
            for i, line in enumerate(skills_lines):
                if start_idx + i < end_idx and start_idx + i < len(doc.paragraphs):
                    para = doc.paragraphs[start_idx + i]
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
    
    # Check if files exist
    if not os.path.exists(CV_PATH):
        return None, None, f"CV file not found: {CV_PATH}"
    
    if not os.path.exists(COVER_PATH):
        return None, None, f"Cover letter template not found: {COVER_PATH}"
    
    # Read documents
    cv_text = read_docx(CV_PATH)
    cover_text = read_docx(COVER_PATH)
    
    if not cv_text or not cover_text:
        return None, None, "Failed to read documents"
    
    # Get tailored content from AI
    try:
        print("🤖 Tailoring CV with DeepSeek...")
        result = tailor_cv_with_ai(cv_text, job_description)
        print(f"✅ CV Tailored: {result.keys() if result else 'None'}")
    except Exception as e:
        return None, None, f"CV tailoring failed: {str(e)}"
    
    # Generate cover letter
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_with_ai(cover_text, cv_text, job_description)
        print(f"✅ Cover letter generated: {len(tailored_cover) if tailored_cover else 0} characters")
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter generation failed, using template: {str(e)}")
    
    # Generate tailored documents with preserved formatting
    try:
        print("📄 Generating tailored CV with formatting preserved...")
        cv_output = preserve_formatting_and_tailor(
            CV_PATH,
            result.get('tailored_summary', ''),
            result.get('tailored_skills', '')
        )
        print("✅ CV generated successfully")
        
        print("📄 Generating cover letter...")
        # For cover letter, replace the entire content
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
        print("✅ Cover letter generated successfully")
        
        return cv_output, cover_output, "Success"
        
    except Exception as e:
        return None, None, f"Document generation error: {str(e)}"

@app.route('/')
def index():
    """Main dashboard page"""
    try:
        return render_template('index.html')
    except Exception as e:
        return f"Error loading template: {str(e)}", 500

@app.route('/tailor', methods=['POST'])
def tailor():
    """Process job description and return tailored documents"""
    try:
        job_description = request.form.get('job_description', '').strip()
        
        if not job_description:
            return jsonify({'error': 'Please paste a job description'}), 400
        
        print(f"📝 Processing job description: {len(job_description)} characters")
        
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
        print(f"❌ Error in /tailor: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    print(f"🚀 Starting CV Tailor on port {port}")
    print(f"📂 Files: {os.listdir('.')}")
    app.run(host='0.0.0.0', port=port, debug=False)
