import os
import json
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from openai import OpenAI
import tempfile

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-in-production")

# Initialize OpenAI
API_KEY = os.environ.get("OPENAI_API_KEY")
if not API_KEY:
    print("⚠️ OPENAI_API_KEY not set. Please set it in environment variables.")
    # For development only - remove in production
    API_KEY = "your-api-key-here"

client = OpenAI(api_key=API_KEY)

# File paths (these will be in the deployment)
CV_PATH = "Master_CV.docx"
COVER_PATH = "Cover_Template.docx"

def read_docx(file_path):
    """Extract text from .docx file"""
    try:
        doc = Document(file_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return ""

def extract_sections(text):
    """Extract sections from CV text"""
    sections = {
        "summary": "",
        "skills": "",
        "experience": "",
        "education": ""
    }
    
    lines = text.split('\n')
    current_section = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        line_lower = line.lower()
        if any(keyword in line_lower for keyword in ['summary', 'profile', 'about']):
            current_section = "summary"
        elif any(keyword in line_lower for keyword in ['skill', 'competencies', 'core']):
            current_section = "skills"
        elif any(keyword in line_lower for keyword in ['experience', 'employment', 'work']):
            current_section = "experience"
        elif any(keyword in line_lower for keyword in ['education', 'academic', 'qualification']):
            current_section = "education"
        elif current_section:
            sections[current_section] += line + "\n"
    
    return sections

def tailor_with_ai(cv_text, cover_text, job_description):
    """Get tailored content from OpenAI"""
    
    # Extract CV sections
    cv_sections = extract_sections(cv_text)
    
    prompt = f"""
You are a professional CV and cover letter tailor.

CV Sections:
Summary: {cv_sections['summary']}
Skills: {cv_sections['skills']}
Experience: {cv_sections['experience']}
Education: {cv_sections['education']}

Cover Letter Template:
{cover_text}

Job Description:
{job_description}

Rules:
1. Rewrite ONLY the Summary and Skills sections to match the job
2. Keep ALL experience, education, and factual information EXACTLY as is
3. For the cover letter, personalize it using the job description
4. Use keywords from the job description naturally
5. NEVER add experience or qualifications not in the CV

Return JSON:
{{
    "tailored_summary": "new summary matching job",
    "tailored_skills": "new skills matching job",
    "tailored_cover": "personalized cover letter"
}}
"""
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an expert CV tailoring assistant."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.3
        )
        
        result = json.loads(response.choices[0].message.content)
        return result
    except Exception as e:
        print(f"OpenAI error: {str(e)}")
        return None

def generate_tailored_docx(template_path, new_text, section_to_replace=None):
    """Generate tailored DOCX preserving formatting"""
    doc = Document(template_path)
    
    if section_to_replace:
        # Replace only specific section
        for paragraph in doc.paragraphs:
            text_lower = paragraph.text.lower()
            if any(keyword in text_lower for keyword in ['summary', 'profile']):
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
            elif any(keyword in text_lower for keyword in ['skill', 'core', 'competencies']):
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
    else:
        # Replace entire document (for cover letter)
        new_paragraphs = [p for p in new_text.split('\n') if p.strip()]
        
        para_index = 0
        for paragraph in doc.paragraphs:
            if para_index >= len(new_paragraphs):
                break
            if paragraph.text.strip() and paragraph.runs:
                paragraph.runs[0].text = new_paragraphs[para_index]
                for run in paragraph.runs[1:]:
                    run.text = ""
                para_index += 1
    
    # Save to BytesIO
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def process_application(job_description):
    """Process the entire application tailoring"""
    
    # Check if files exist
    if not os.path.exists(CV_PATH):
        return None, None, "CV file not found"
    
    if not os.path.exists(COVER_PATH):
        return None, None, "Cover letter template not found"
    
    # Read documents
    cv_text = read_docx(CV_PATH)
    cover_text = read_docx(COVER_PATH)
    
    if not cv_text or not cover_text:
        return None, None, "Failed to read documents"
    
    # Get tailored content from AI
    result = tailor_with_ai(cv_text, cover_text, job_description)
    
    if not result:
        return None, None, "AI processing failed"
    
    # Generate tailored documents
    try:
        # Tailor CV - replace only summary and skills
        cv_doc = Document(CV_PATH)
        
        # Update summary
        for paragraph in cv_doc.paragraphs:
            text_lower = paragraph.text.lower()
            if any(keyword in text_lower for keyword in ['summary', 'profile', 'about']):
                if paragraph.runs:
                    paragraph.runs[0].text = result.get('tailored_summary', paragraph.text)
        
        # Update skills
        for paragraph in cv_doc.paragraphs:
            text_lower = paragraph.text.lower()
            if any(keyword in text_lower for keyword in ['skill', 'core', 'competencies']):
                if paragraph.runs:
                    paragraph.runs[0].text = result.get('tailored_skills', paragraph.text)
        
        cv_output = BytesIO()
        cv_doc.save(cv_output)
        cv_output.seek(0)
        
        # Generate cover letter
        cover_doc = Document(COVER_PATH)
        cover_new_text = result.get('tailored_cover', '')
        
        if cover_new_text:
            new_paragraphs = [p for p in cover_new_text.split('\n') if p.strip()]
            para_index = 0
            for paragraph in cover_doc.paragraphs:
                if para_index >= len(new_paragraphs):
                    break
                if paragraph.text.strip() and paragraph.runs:
                    paragraph.runs[0].text = new_paragraphs[para_index]
                    para_index += 1
        
        cover_output = BytesIO()
        cover_doc.save(cover_output)
        cover_output.seek(0)
        
        return cv_output, cover_output, "Success"
        
    except Exception as e:
        return None, None, f"Document generation error: {str(e)}"

# Routes
@app.route('/')
def index():
    """Main dashboard page"""
    return render_template('index.html')

@app.route('/tailor', methods=['POST'])
def tailor():
    """Process job description and return tailored documents"""
    try:
        job_description = request.form.get('job_description', '').strip()
        
        if not job_description:
            return jsonify({'error': 'Please paste a job description'}), 400
        
        # Process the application
        cv_output, cover_output, message = process_application(job_description)
        
        if cv_output and cover_output:
            # Generate download links
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            return jsonify({
                'success': True,
                'message': 'Documents tailored successfully!',
                'cv_filename': f'Tailored_CV_{timestamp}.docx',
                'cover_filename': f'Tailored_Cover_{timestamp}.docx',
                'cv_data': cv_output.getvalue().decode('latin1'),  # Convert to base64
                'cover_data': cover_output.getvalue().decode('latin1')
            })
        else:
            return jsonify({'error': message}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/cv')
def download_cv():
    """Download tailored CV"""
    try:
        # For demonstration - in production, you'd store the file
        cv_output, _, _ = process_application(request.args.get('job_description', ''))
        if cv_output:
            return send_file(
                cv_output,
                as_attachment=True,
                download_name=f"Tailored_CV_{datetime.now().strftime('%Y%m%d')}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        return "CV not found", 404
    except Exception as e:
        return str(e), 500

@app.route('/download/cover')
def download_cover():
    """Download tailored cover letter"""
    try:
        _, cover_output, _ = process_application(request.args.get('job_description', ''))
        if cover_output:
            return send_file(
                cover_output,
                as_attachment=True,
                download_name=f"Tailored_Cover_{datetime.now().strftime('%Y%m%d')}.docx",
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        return "Cover letter not found", 404
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
