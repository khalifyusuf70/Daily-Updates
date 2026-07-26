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

# ---------------------------
# AI PROMPTS (From Your Gist)
# ---------------------------
SYSTEM_PROMPT = """
You are an expert executive resume writer and ATS optimization specialist.
Your job is to tailor an existing CV to a target job description.

Rules:
- NEVER invent experience, achievements, qualifications, employers, or certifications.
- Rewrite existing experience to better match the job description.
- Prioritize keywords naturally.
- Improve the Professional Summary.
- Rewrite bullet points using the user's actual experience.
- Optimize for ATS while remaining truthful.
- Return only the new summary and updated experience bullets.
"""

def build_user_prompt(cv_text, job_description):
    return f"""
MASTER CV
{cv_text}

JOB DESCRIPTION
{job_description}

Tasks:
1. Rewrite the Professional Summary to closely match the job description.
2. Rewrite each job description under Experience so it aligns with the advertised role while remaining truthful.
3. Add relevant ATS keywords where appropriate.
4. Do not fabricate any experience.

Return JSON in this format:
{{
    "summary": "...",
    "experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "...",
            "...",
            "..."
        ],
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": [
            "...",
            "..."
        ],
        "Chief Operations Officer | Jul 2016 -- Jul 2021": [
            "...",
            "..."
        ]
    }}
}}
"""

def call_ai(cv_text, job_description):
    """Call DeepSeek API with the tailored prompts"""
    try:
        response = client.chat.completions.create(
            model="deepseek-v4-pro",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(cv_text, job_description)},
            ],
            temperature=0.2,
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"DeepSeek API error: {str(e)}")
        raise e

# ---------------------------
# HELPER FUNCTIONS
# ---------------------------
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

def extract_job_titles(cv_text):
    """Extract exact job titles from CV text for matching"""
    job_titles = []
    lines = cv_text.split('\n')
    for line in lines:
        if '(' in line and ')' in line and any(y in line for y in ['2023', '2022', '2021', '2020', '2019', '2018', '2017', '2016']):
            if any(keyword in line.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager', 'director']):
                clean = re.sub(r'\{#.*?\}', '', line)
                clean = re.sub(r'\.Styl\d+', '', clean)
                clean = re.sub(r'#', '', line)
                if clean.strip():
                    job_titles.append(clean.strip())
    return job_titles

def tailor_cover_letter(cover_text, cv_text, job_description):
    """Generate tailored cover letter"""
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
            model="deepseek-v4-pro",
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

def create_tailored_docx(template_path, new_summary, new_experience):
    """
    Create tailored CV with ALL formatting requirements:
    - Calibri Body 12 for all text
    - Segoe UI 12 Bold for job titles
    - Segoe UI 12 Italic for company names
    - Skills in two-column table format
    - Experience in correct order (most recent first)
    """
    doc = Document(template_path)
    
    # Find section positions
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
    
    print(f"\n📌 Found: Summary={summary_pos}, Skills={skills_pos}, Experience={experience_pos}")
    
    # 1. UPDATE SUMMARY - Calibri Body 12
    if summary_pos != -1 and new_summary:
        end_pos = skills_pos if skills_pos > summary_pos else len(doc.paragraphs)
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        if summary_pos + 1 < len(doc.paragraphs):
            para = doc.paragraphs[summary_pos + 1]
            para.text = new_summary
            # Apply Calibri Body 12
            for run in para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
        print(f"✅ Summary updated with Calibri 12")
    
    # 2. UPDATE SKILLS - Two-column table format
    if skills_pos != -1:
        print(f"\n📝 Rebuilding skills section...")
        
        # Define skills in two columns as requested
        skills_left = [
            "• Resource Mobilization & Grants Management",
            "• Foundation Pipeline Management",
            "• Due Diligence & Risk Assessment",
            "• Impact Reporting & Communication",
            "• Political & Contextual Analysis"
        ]
        skills_right = [
            "• Team Leadership & Organizational Development",
            "• Risk Management & Compliance",
            "• Humanitarian Response Programming",
            "• Strategic Planning & Organizational Dev."
        ]
        
        # Find end of skills section
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        
        # Clear existing skills content
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Create a table for skills (2 columns)
        # Find the position to insert table
        table_pos = skills_pos + 1
        
        # Remove any existing table at this position
        for i in range(table_pos, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Create new table with 2 columns
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        table.autofit = False
        
        # Set column widths
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(3.5)
        
        # Fill table with skills
        for row_idx in range(5):
            # Left column
            left_cell = table.cell(row_idx, 0)
            left_cell.text = skills_left[row_idx] if row_idx < len(skills_left) else ""
            left_para = left_cell.paragraphs[0]
            for run in left_para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
            
            # Right column
            right_cell = table.cell(row_idx, 1)
            right_cell.text = skills_right[row_idx] if row_idx < len(skills_right) else ""
            right_para = right_cell.paragraphs[0]
            for run in right_para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
        
        # Move table to correct position
        # Get the table element and move it after the header
        table_element = table._element
        parent = doc.element.body
        # Insert table after the skills header
        parent.insert(summary_pos + 2, table_element)
        
        print(f"✅ Skills table created with Calibri 12")
    
    # 3. UPDATE EXPERIENCE - With proper formatting
    if experience_pos != -1 and new_experience:
        print(f"\n📝 Updating experience with formatting...")
        
        # Get job titles in correct order (most recent first)
        job_order = [
            "Chief of Staff (Feb 2023-To Date)",
            "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023",
            "Chief Operations Officer | Jul 2016 -- Jul 2021"
        ]
        
        # Find job positions in document
        job_positions = []
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if '(' in text and ')' in text:
                if any(keyword in text.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager']):
                    job_positions.append((i, text))
        
        # Process each job in order
        for job_title in job_order:
            if job_title not in new_experience:
                continue
            
            new_bullets = new_experience[job_title]
            print(f"  Looking for: {job_title[:40]}...")
            
            # Find matching job
            job_idx = -1
            for i, text in job_positions:
                if job_title in text or text in job_title:
                    job_idx = i
                    break
            
            if job_idx == -1:
                print(f"    ⚠️ Job not found: {job_title[:40]}")
                continue
            
            print(f"    ✅ Found at index {job_idx}")
            
            # Find end of this job section
            end_idx = len(doc.paragraphs)
            for next_pos, _ in job_positions:
                if next_pos > job_idx and next_pos < end_idx:
                    end_idx = next_pos
                    break
            
            # Find employer line (company name)
            company_name = ""
            for i in range(job_idx + 1, end_idx):
                if i < len(doc.paragraphs):
                    text = doc.paragraphs[i].text.strip()
                    if any(keyword in text for keyword in ['Jubaland', 'Ministry', 'KIMS']):
                        company_name = text
                        break
            
            # Find bullet points
            bullet_indices = []
            for i in range(job_idx + 1, end_idx):
                if i < len(doc.paragraphs):
                    text = doc.paragraphs[i].text.strip()
                    if text.startswith('-') or text.startswith('•') or text.startswith('*'):
                        bullet_indices.append(i)
            
            # Clear existing bullets
            for i in bullet_indices:
                if i < len(doc.paragraphs):
                    doc.paragraphs[i].text = ""
            
            # Insert new bullets
            for i, bullet_text in enumerate(new_bullets):
                insert_pos = job_idx + 2 + i  # After job title and company
                if insert_pos < len(doc.paragraphs):
                    para = doc.paragraphs[insert_pos]
                    para.text = f"• {bullet_text}"
                    for run in para.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(12)
                else:
                    # Add new paragraph
                    para = doc.add_paragraph(f"• {bullet_text}")
                    for run in para.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(12)
            
            # Apply formatting to job title (Segoe UI 12 Bold)
            if job_idx < len(doc.paragraphs):
                para = doc.paragraphs[job_idx]
                for run in para.runs:
                    run.font.name = 'Segoe UI'
                    run.font.size = Pt(12)
                    run.font.bold = True
            
            # Apply formatting to company name (Segoe UI 12 Italic)
            if company_name:
                for i in range(job_idx + 1, end_idx):
                    if i < len(doc.paragraphs):
                        text = doc.paragraphs[i].text.strip()
                        if any(keyword in text for keyword in ['Jubaland', 'Ministry', 'KIMS']):
                            para = doc.paragraphs[i]
                            for run in para.runs:
                                run.font.name = 'Segoe UI'
                                run.font.size = Pt(12)
                                run.font.italic = True
                            break
            
            print(f"    ✅ Updated with {len(new_bullets)} bullets")
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ---------------------------
# MAIN PROCESSING
# ---------------------------
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
        print("🤖 Tailoring CV with AI...")
        result = call_ai(cv_text, job_description)
        
        tailored_summary = result.get('summary', '')
        tailored_experience = result.get('experience', {})
        
        print(f"📝 Summary: {tailored_summary[:100]}...")
        print(f"📝 Jobs: {list(tailored_experience.keys())}")
        
    except Exception as e:
        return None, None, f"AI tailoring failed: {str(e)}"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter(cover_text, cv_text, job_description)
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter failed: {str(e)}")
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = create_tailored_docx(CV_PATH, tailored_summary, tailored_experience)
        
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

# ---------------------------
# ROUTES
# ---------------------------
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
