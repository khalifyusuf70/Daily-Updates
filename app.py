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
# STORED MASTER EXPERIENCES
# ---------------------------
MASTER_EXPERIENCE = {
    "Chief of Staff (Feb 2023-To Date)": {
        "company": "Jubaland State of Somalia",
        "bullets": [
            "Directed the Office of the President, coordinating activities across 15 government ministries and agencies to ensure strategic alignment with the State Development Plan, mirroring cross-functional team coordination for communication campaigns.",
            "Represented the President at national and regional forums, managing relationships with international donors (USAID, EU, UN), diplomatic missions, and development partners to advance policy and program alignment.",
            "Spearheaded emergency response efforts during political and humanitarian crises, demonstrating rapid adaptation to evolving contexts and ensuring coordinated resource allocation and stakeholder communication.",
            "Mobilized donor funding and cultivated strategic partnerships with international NGOs and UN agencies to support governance, service delivery, and advocacy initiatives, enhancing program sustainability."
        ]
    },
    "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": {
        "company": "Ministry of Planning, Jubaland State, Somalia",
        "bullets": [
            "Led humanitarian and development needs assessments, identifying priority program areas and in-country stakeholders to inform strategic planning and successful funding proposals for international donors.",
            "Managed a portfolio of donor-funded programs, monitoring deliverables, budgets, and compliance to ensure alignment with grant agreements and reporting deadlines.",
            "Established and maintained partnerships with UN agencies, international NGOs, and civil society organizations to foster collaboration on peacebuilding and advocacy initiatives, strengthening regional outreach.",
            "Conducted political, economic, and social analysis to provide data-driven insights for ministry strategy, keeping abreast of country contexts to support informed decision-making and external communications."
        ]
    },
    "Chief Operations Officer | Jul 2016 -- Jul 2021": {
        "company": "KIMS MICROFINANCE, Somalia",
        "bullets": [
            "Directed all operational functions, including budgeting, strategic planning, and resource allocation, to drive organizational development and ensure efficient program delivery.",
            "Led business development initiatives, expanding into new regions and securing funding for microenterprise programs that supported vulnerable populations, demonstrating multi-regional grant management capability.",
            "Established and managed strategic partnerships with international development organizations to enhance program impact and sustainability, mirroring the coordination of local grantees and partners.",
            "Represented the organization to government agencies, donors, and partners, managing stakeholder relations and contributing to policy discussions on economic empowerment and social development."
        ]
    }
}

JOB_ORDER = [
    "Chief of Staff (Feb 2023-To Date)",
    "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023",
    "Chief Operations Officer | Jul 2016 -- Jul 2021"
]

# ---------------------------
# AI PROMPTS
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
    exact_titles = list(MASTER_EXPERIENCE.keys())
    titles_json = json.dumps(exact_titles, indent=2)
    
    # Build template
    exp_template = {}
    for title in exact_titles:
        exp_template[title] = ["bullet 1", "bullet 2"]
    exp_json = json.dumps(exp_template, indent=2)
    
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

CRITICAL: Use EXACTLY these job titles as keys (copy them exactly):
{titles_json}

Return JSON in this format:
{{
    "summary": "...",
    "experience": {exp_json}
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

def build_cv_text_from_master():
    """Build CV text from stored master experience"""
    cv_text = ""
    for job_title in JOB_ORDER:
        if job_title in MASTER_EXPERIENCE:
            job_data = MASTER_EXPERIENCE[job_title]
            cv_text += f"{job_title}\n"
            cv_text += f"{job_data['company']}\n"
            for bullet in job_data['bullets']:
                cv_text += f"- {bullet}\n"
            cv_text += "\n"
    return cv_text

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
    Create tailored CV with ALL formatting requirements
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
    
    # 1. UPDATE SUMMARY
    if summary_pos != -1 and new_summary:
        end_pos = skills_pos if skills_pos > summary_pos else len(doc.paragraphs)
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        if summary_pos + 1 < len(doc.paragraphs):
            para = doc.paragraphs[summary_pos + 1]
            para.text = new_summary
            for run in para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
        print(f"✅ Summary updated with Calibri 12")
    
    # 2. UPDATE SKILLS
    if skills_pos != -1:
        print(f"\n📝 Rebuilding skills section...")
        
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
        
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        table = doc.add_table(rows=5, cols=2)
        table.style = 'Table Grid'
        table.autofit = False
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(3.5)
        
        for row_idx in range(5):
            left_cell = table.cell(row_idx, 0)
            left_cell.text = skills_left[row_idx] if row_idx < len(skills_left) else ""
            left_para = left_cell.paragraphs[0]
            for run in left_para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
            
            right_cell = table.cell(row_idx, 1)
            right_cell.text = skills_right[row_idx] if row_idx < len(skills_right) else ""
            right_para = right_cell.paragraphs[0]
            for run in right_para.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(12)
        
        table_element = table._element
        parent = doc.element.body
        parent.insert(skills_pos + 2, table_element)
        
        print(f"✅ Skills table created with Calibri 12")
    
    # 3. UPDATE EXPERIENCE with improved matching
    if experience_pos != -1 and new_experience:
        print(f"\n📝 Updating experience with formatting...")
        
        # Get ALL job titles from the document
        doc_job_titles = []
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if '(' in text and ')' in text:
                if any(keyword in text.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager']):
                    doc_job_titles.append((i, text))
        
        print(f"📌 Document job titles found: {len(doc_job_titles)}")
        for i, title in doc_job_titles:
            print(f"  - {title[:50]}...")
        
        for job_title, new_bullets in new_experience.items():
            print(f"  Looking for: {job_title[:40]}...")
            
            job_idx = -1
            matched_title = ""
            
            # Strategy 1: Exact match
            for i, text in doc_job_titles:
                if text == job_title:
                    job_idx = i
                    matched_title = text
                    break
            
            # Strategy 2: Contains match
            if job_idx == -1:
                for i, text in doc_job_titles:
                    if job_title in text or text in job_title:
                        job_idx = i
                        matched_title = text
                        break
            
            # Strategy 3: Clean and compare
            if job_idx == -1:
                clean_job = re.sub(r'[^a-zA-Z0-9 ]', '', job_title.lower())
                for i, text in doc_job_titles:
                    clean_text = re.sub(r'[^a-zA-Z0-9 ]', '', text.lower())
                    if clean_job in clean_text or clean_text in clean_job:
                        job_idx = i
                        matched_title = text
                        break
            
            # Strategy 4: Keyword matching
            if job_idx == -1:
                keywords = re.findall(r'\b(Chief|Senior|Advisor|Officer|Manager)\b', job_title, re.IGNORECASE)
                for i, text in doc_job_titles:
                    if any(keyword.lower() in text.lower() for keyword in keywords):
                        if any(year in text for year in ['2023', '2022', '2021', '2020']):
                            job_idx = i
                            matched_title = text
                            break
            
            if job_idx == -1:
                print(f"    ⚠️ Job not found: {job_title[:40]}")
                continue
            
            print(f"    ✅ Found at index {job_idx}: {matched_title[:40]}")
            
            # Find end of this job section
            end_idx = len(doc.paragraphs)
            for next_pos, _ in doc_job_titles:
                if next_pos > job_idx and next_pos < end_idx:
                    end_idx = next_pos
                    break
            
            # Update company name formatting
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
            
            # Find bullet points
            bullet_indices = []
            for i in range(job_idx + 1, end_idx):
                if i < len(doc.paragraphs):
                    text = doc.paragraphs[i].text.strip()
                    if text.startswith('-') or text.startswith('•') or text.startswith('*'):
                        bullet_indices.append(i)
            
            print(f"    Found {len(bullet_indices)} bullet points")
            
            # Clear existing bullets
            for i in bullet_indices:
                if i < len(doc.paragraphs):
                    doc.paragraphs[i].text = ""
            
            # Insert new bullets
            for i, bullet_text in enumerate(new_bullets):
                insert_pos = job_idx + 2 + i
                if insert_pos < len(doc.paragraphs):
                    para = doc.paragraphs[insert_pos]
                    para.text = f"• {bullet_text}"
                    for run in para.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(12)
                else:
                    para = doc.add_paragraph(f"• {bullet_text}")
                    for run in para.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(12)
            
            # Apply formatting to job title
            if job_idx < len(doc.paragraphs):
                para = doc.paragraphs[job_idx]
                for run in para.runs:
                    run.font.name = 'Segoe UI'
                    run.font.size = Pt(12)
                    run.font.bold = True
            
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
    
    cv_text = build_cv_text_from_master()
    cover_text = read_docx(COVER_PATH)
    
    if not cv_text or not cover_text:
        return None, None, "Failed to read documents"
    
    print(f"📄 CV text built from master experience: {len(cv_text)} characters")
    
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
