import os
import sys
import json
import base64
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-in-production")

# ---------------------------
# DEEPSEEK API CONFIGURATION
# ---------------------------
try:
    from openai import OpenAI
except ImportError as e:
    print(f"Error importing OpenAI: {e}")
    sys.exit(1)

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
if not DEEPSEEK_API_KEY:
    print("⚠️ DEEPSEEK_API_KEY not set. Please set it in environment variables.")
    # For demo purposes only - exit in production
    sys.exit(1)

# Initialize DeepSeek client
try:
    client = OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url="https://api.deepseek.com/v1"
    )
    print("✅ DeepSeek client initialized successfully")
except Exception as e:
    print(f"❌ Failed to initialize DeepSeek client: {e}")
    sys.exit(1)

# File paths
CV_PATH = "Master_CV.docx"
COVER_PATH = "Cover_Template.docx"

# Check if files exist
print(f"📂 Checking for CV file: {CV_PATH}")
if not os.path.exists(CV_PATH):
    print(f"❌ CV file not found: {CV_PATH}")
    print(f"📂 Directory contents: {os.listdir('.')}")

print(f"📂 Checking for Cover file: {COVER_PATH}")
if not os.path.exists(COVER_PATH):
    print(f"❌ Cover file not found: {COVER_PATH}")

def call_deepseek(prompt):
    """Call DeepSeek API with proper format"""
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert CV tailoring assistant."},
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
    """Get tailored content from DeepSeek"""
    
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
    
    return call_deepseek(prompt)

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
        result = tailor_with_ai(cv_text, cover_text, job_description)
    except Exception as e:
        return None, None, f"AI processing failed: {str(e)}"
    
    if not result:
        return None, None, "AI returned empty result"
    
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
        
        # Process the application
        cv_output, cover_output, message = process_application(job_description)
        
        if cv_output and cover_output:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Convert to base64 for safe JSON transmission
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
        print(f"Error in /tailor: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    """Health check endpoint for Railway"""
    return jsonify({'status': 'healthy', 'timestamp': datetime.now().isoformat()})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    print(f"🚀 Starting Flask app on port {port}")
    print(f"📂 Current directory: {os.getcwd()}")
    print(f"📂 Files: {os.listdir('.')}")
    app.run(host='0.0.0.0', port=port, debug=False)
