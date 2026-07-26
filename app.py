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

# Master CV template - this will be used to build the document
MASTER_CV_TEMPLATE = """
# Summary
[Summary will be replaced by AI]

# Skill Highlights
[Skills will be inserted here]

# Experience

Chief of Staff (Feb 2023-To Date)
Jubaland State of Somalia
- Directed the Office of the President, coordinating activities across 15 government ministries and agencies to ensure strategic alignment with the State Development Plan, mirroring cross-functional team coordination for communication campaigns.
- Represented the President at national and regional forums, managing relationships with international donors (USAID, EU, UN), diplomatic missions, and development partners to advance policy and program alignment.
- Spearheaded emergency response efforts during political and humanitarian crises, demonstrating rapid adaptation to evolving contexts and ensuring coordinated resource allocation and stakeholder communication.
- Mobilized donor funding and cultivated strategic partnerships with international NGOs and UN agencies to support governance, service delivery, and advocacy initiatives, enhancing program sustainability.

Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023
Ministry of Planning, Jubaland State, Somalia
- Led humanitarian and development needs assessments, identifying priority program areas and in-country stakeholders to inform strategic planning and successful funding proposals for international donors.
- Managed a portfolio of donor-funded programs, monitoring deliverables, budgets, and compliance to ensure alignment with grant agreements and reporting deadlines.
- Established and maintained partnerships with UN agencies, international NGOs, and civil society organizations to foster collaboration on peacebuilding and advocacy initiatives, strengthening regional outreach.
- Conducted political, economic, and social analysis to provide data-driven insights for ministry strategy, keeping abreast of country contexts to support informed decision-making and external communications.

Chief Operations Officer | Jul 2016 -- Jul 2021
KIMS MICROFINANCE, Somalia
- Directed all operational functions, including budgeting, strategic planning, and resource allocation, to drive organizational development and ensure efficient program delivery.
- Led business development initiatives, expanding into new regions and securing funding for microenterprise programs that supported vulnerable populations, demonstrating multi-regional grant management capability.
- Established and managed strategic partnerships with international development organizations to enhance program impact and sustainability, mirroring the coordination of local grantees and partners.
- Represented the organization to government agencies, donors, and partners, managing stakeholder relations and contributing to policy discussions on economic empowerment and social development.

# Education
[Your education here]

# REFREES
[Your referees here]
"""

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
        "Chief of Staff (Feb 2023-To Date)": ["...", "..."],
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": ["...", "..."],
        "Chief Operations Officer | Jul 2016 -- Jul 2021": ["...", "..."]
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
# COVER LETTER
# ---------------------------
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

def create_tailored_docx(new_summary, new_experience):
    """
    Create tailored CV from scratch with ALL formatting requirements
    """
    # Create a new document
    doc = Document()
    
    # Set default font to Calibri 12
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(12)
    
    # Title/Header
    header_para = doc.add_paragraph("+252611385559/+254715244144")
    header_para.runs[0].font.name = 'Calibri'
    header_para.runs[0].font.size = Pt(12)
    
    doc.add_paragraph("cos.presidency@jubalandstate.so")
    doc.add_paragraph("khalifyusuf@agileanalytica.com")
    doc.add_paragraph("")
    
    # 1. SUMMARY SECTION
    summary_header = doc.add_paragraph("# Summary")
    summary_header.runs[0].font.name = 'Calibri'
    summary_header.runs[0].font.size = Pt(12)
    summary_header.runs[0].bold = True
    
    if new_summary:
        summary_para = doc.add_paragraph(new_summary)
        summary_para.runs[0].font.name = 'Calibri'
        summary_para.runs[0].font.size = Pt(12)
    doc.add_paragraph("")
    
    # 2. SKILLS SECTION
    skills_header = doc.add_paragraph("# Skill Highlights")
    skills_header.runs[0].font.name = 'Calibri'
    skills_header.runs[0].font.size = Pt(12)
    skills_header.runs[0].bold = True
    
    # Skills table (2 columns)
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
    
    doc.add_paragraph("")
    
    # 3. EXPERIENCE SECTION
    exp_header = doc.add_paragraph("# Experience")
    exp_header.runs[0].font.name = 'Calibri'
    exp_header.runs[0].font.size = Pt(12)
    exp_header.runs[0].bold = True
    
    # Job order
    JOB_ORDER = [
        "Chief of Staff (Feb 2023-To Date)",
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023",
        "Chief Operations Officer | Jul 2016 -- Jul 2021"
    ]
    
    # Company names
    companies = {
        "Chief of Staff (Feb 2023-To Date)": "Jubaland State of Somalia",
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": "Ministry of Planning, Jubaland State, Somalia",
        "Chief Operations Officer | Jul 2016 -- Jul 2021": "KIMS MICROFINANCE, Somalia"
    }
    
    for job_title in JOB_ORDER:
        # Job Title - Segoe UI 12 Bold
        job_para = doc.add_paragraph(job_title)
        job_para.runs[0].font.name = 'Segoe UI'
        job_para.runs[0].font.size = Pt(12)
        job_para.runs[0].bold = True
        
        # Company - Segoe UI 12 Italic
        company_para = doc.add_paragraph(companies.get(job_title, ""))
        company_para.runs[0].font.name = 'Segoe UI'
        company_para.runs[0].font.size = Pt(12)
        company_para.runs[0].italic = True
        
        # Bullets
        if job_title in new_experience:
            for bullet in new_experience[job_title]:
                bullet_para = doc.add_paragraph(f"• {bullet}")
                bullet_para.runs[0].font.name = 'Calibri'
                bullet_para.runs[0].font.size = Pt(12)
        else:
            # Fallback to master bullets
            master_bullets = {
                "Chief of Staff (Feb 2023-To Date)": [
                    "Directed the Office of the President, coordinating activities across 15 government ministries and agencies to ensure strategic alignment with the State Development Plan.",
                    "Represented the President at national and regional forums, managing relationships with international donors (USAID, EU, UN), diplomatic missions, and development partners.",
                    "Spearheaded emergency response efforts during political and humanitarian crises, ensuring coordinated resource allocation and stakeholder communication.",
                    "Mobilized donor funding and cultivated strategic partnerships with international NGOs and UN agencies to support governance and service delivery."
                ],
                "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": [
                    "Led humanitarian and development needs assessments to inform strategic planning and funding proposals for international donors.",
                    "Managed a portfolio of donor-funded programs, monitoring deliverables, budgets, and compliance.",
                    "Established partnerships with UN agencies, NGOs, and civil society organizations to foster collaboration.",
                    "Conducted political, economic, and social analysis to provide data-driven insights for ministry strategy."
                ],
                "Chief Operations Officer | Jul 2016 -- Jul 2021": [
                    "Directed all operational functions including budgeting, strategic planning, and resource allocation.",
                    "Led business development initiatives, expanding into new regions and securing funding for microenterprise programs.",
                    "Established strategic partnerships with international development organizations to enhance program impact.",
                    "Represented the organization to government agencies, donors, and partners."
                ]
            }
            for bullet in master_bullets.get(job_title, []):
                bullet_para = doc.add_paragraph(f"• {bullet}")
                bullet_para.runs[0].font.name = 'Calibri'
                bullet_para.runs[0].font.size = Pt(12)
        
        doc.add_paragraph("")
    
    # 4. EDUCATION SECTION
    edu_header = doc.add_paragraph("# Education")
    edu_header.runs[0].font.name = 'Calibri'
    edu_header.runs[0].font.size = Pt(12)
    edu_header.runs[0].bold = True
    
    # Education from original
    education = [
        "Micro Masters. Database management System",
        "University of Baltimore Maryland",
        "Degree Actuarial Science and Statistics",
        "JOMO KENYATTA UNIVERSITY",
        "Nairobi, Kenya"
    ]
    for edu in education:
        edu_para = doc.add_paragraph(edu)
        edu_para.runs[0].font.name = 'Calibri'
        edu_para.runs[0].font.size = Pt(12)
    
    doc.add_paragraph("")
    
    # 5. REFREES SECTION
    ref_header = doc.add_paragraph("# REFREES")
    ref_header.runs[0].font.name = 'Calibri'
    ref_header.runs[0].font.size = Pt(12)
    ref_header.runs[0].bold = True
    
    referees = [
        "Abshir olow",
        "Former Chief of Staff",
        "Jubaland State, Somalia",
        "+254722877178",
        "abshirmabdi@gmail.com",
        "",
        "Mr. Abdirahman Abdi",
        "Planning Minister, Jubaland",
        "+252612992848",
        "mopic@jubalandstate.so",
        "",
        "Omar Abdi Mohamed, HDmls, Msc.Int.",
        "Commodity Logistics Director-USAID",
        "omarabdi2@gmail.com"
    ]
    for ref in referees:
        if ref:
            ref_para = doc.add_paragraph(ref)
            ref_para.runs[0].font.name = 'Calibri'
            ref_para.runs[0].font.size = Pt(12)
        else:
            doc.add_paragraph()
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ---------------------------
# MAIN PROCESSING
# ---------------------------
def process_application(job_description):
    """Process the entire application tailoring"""
    
    # Use the template as CV text
    cv_text = MASTER_CV_TEMPLATE
    
    # Get cover letter template
    if os.path.exists(COVER_PATH):
        cover_text = read_docx(COVER_PATH)
    else:
        cover_text = "Dear Hiring Manager,\n\n[Your cover letter will go here]\n\nSincerely,\n[Your Name]"
    
    if not cv_text or not cover_text:
        return None, None, "Failed to read documents"
    
    print(f"📄 CV text built from master template: {len(cv_text)} characters")
    
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
        cv_output = create_tailored_docx(tailored_summary, tailored_experience)
        
        # Cover letter
        cover_doc = Document(COVER_PATH) if os.path.exists(COVER_PATH) else Document()
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
