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
import traceback

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
    base_url="https://api.deepseek.com"
)

COVER_PATH = "Cover_Template.docx"

# ---------------------------
# MASTER CV TEMPLATE
# ---------------------------
MASTER_CV_TEMPLATE = """
Chief of Staff (Feb 2023-To Date) | Jubaland State of Somalia
- Donor Engagement & Fundraising: Led resource mobilization efforts, engaging with international donors (USAID, EU, UN), diplomatic missions, and development partners to secure funding for state priorities and emergency response.
- Strategic Partnership Development: Established and maintained partnerships with international NGOs, UN agencies, and civil society organizations to enhance governance, humanitarian response, and service delivery.
- Strategic Leadership: Directed the Office of the President, coordinating activities across 15 government ministries and agencies to ensure alignment with the State Development Plan.
- Program Growth & Strategy: Led the development and implementation of the comprehensive State Development Plan, translating high-level goals into actionable programs and policies.
- Stakeholder Management: Represented the President at national and regional forums, managing complex relationships and ensuring aligned engagement across all partners.
- Crisis Management: Spearheaded emergency response efforts during political instability and humanitarian crises, coordinating relief activities and resource allocation.
- Established and maintained partnerships with civil society organizations and community leaders to enhance governance and service delivery.

Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023) | Ministry of Planning, Jubaland State, Somalia
- Grant Acquisition & Management: Led the development of the Ministry's Strategic Plan and managed a portfolio of donor-funded programs, ensuring effective implementation, financial accountability, and compliance.
- Proposal Development: Led humanitarian and development needs assessments, translating findings into priority program areas and successful funding proposals for international donors.
- Strategy & Policy: Provided strategic leadership, advising the Minister on planning, policy, and program development to align Ministry initiatives with national and regional priorities.
- External Relations: Established and maintained partnerships with international development organizations, UN agencies, and NGOs, fostering collaboration on civil society strengthening and peacebuilding initiatives.
- Analytical Reporting: Conducted political, economic, and social analysis to inform Ministry strategy, providing critical data-driven insights for programmatic decision-making.
- Capacity Building: Strengthened institutional capacity through training and technical support to ministry staff and partners.

Chief Operations Officer (Jul 2016 – Jul 2021) | KIMS MICROFINANCE, Somalia
- Organizational Leadership: Directed all operational functions, including budgeting, strategic planning, and resource allocation, to drive organizational development and growth.
- Business Development: Led business development initiatives, expanding the organization's reach into new regions and securing funding for microenterprise programming that supported vulnerable populations.
- Partnership Management: Established and maintained strategic partnerships with international development organizations to enhance program impact and sustainability.
- Community Engagement: Managed community relations, contributing to peacebuilding and social development through economic empowerment and locally-led programming.
- Stakeholder Relations: Represented the organization to government agencies, donors, and partners, ensuring effective communication and stakeholder management.
- Financial Oversight: Oversaw financial management, compliance, and reporting to ensure transparency and accountability.
- Team Leadership: Built and led high-performing teams, fostering a culture of excellence and continuous improvement.
"""

# ---------------------------
# EXACT JOB TITLES
# ---------------------------
JOB_TITLES = [
    "Chief of Staff (Feb 2023-To Date)",
    "Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)",
    "Chief Operations Officer (Jul 2016 – Jul 2021)"
]

COMPANY_NAMES = {
    "Chief of Staff (Feb 2023-To Date)": "Jubaland State of Somalia",
    "Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)": "Ministry of Planning, Jubaland State, Somalia",
    "Chief Operations Officer (Jul 2016 – Jul 2021)": "KIMS MICROFINANCE, Somalia"
}

# ================================================================
# IMPROVED PROMPTS - THIS IS THE KEY DIFFERENCE!
# ================================================================

SYSTEM_PROMPT = """
You are one of the world's leading Executive Resume Writers, Executive Recruiters,
ATS Optimization Specialists, and Career Strategists.

Your task is to tailor a candidate's CV and cover letter to maximize alignment
with a target job description while maintaining complete honesty and factual accuracy.

Rules:
- Never fabricate employers, education, achievements or qualifications.
- Rewrite only using truthful information.
- Optimize for ATS (Applicant Tracking Systems).
- Rewrite the Professional Summary.
- Rewrite Core Competencies.
- Rewrite Experience sections with relevant keywords.
- Write a tailored Cover Letter.
"""

def build_user_prompt(cv_text, job_description):
    return f"""
MASTER CV
{cv_text}

JOB DESCRIPTION
{job_description}

Carefully analyze both documents.

Rewrite the CV specifically for this position.

Return JSON with:
1. Updated Professional Summary
2. Updated Core Competencies
3. Updated Experience (for each job, with 6-7 bullet points)
4. Updated Technical Skills

Return ONLY valid JSON with this format:
{{
    "summary": "4-6 sentence professional summary",
    "core_competencies": ["competency 1", "competency 2", "competency 3", "competency 4", "competency 5"],
    "experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "bullet 1",
            "bullet 2",
            "bullet 3",
            "bullet 4",
            "bullet 5",
            "bullet 6"
        ],
        "Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)": [
            "bullet 1",
            "bullet 2",
            "bullet 3",
            "bullet 4",
            "bullet 5",
            "bullet 6"
        ],
        "Chief Operations Officer (Jul 2016 – Jul 2021)": [
            "bullet 1",
            "bullet 2",
            "bullet 3",
            "bullet 4",
            "bullet 5",
            "bullet 6"
        ]
    }},
    "technical_skills": ["skill 1", "skill 2", "skill 3", "skill 4", "skill 5"]
}}
"""

def build_cover_letter_prompt(cv_text, job_description):
    return f"""
MASTER CV:
{cv_text}

JOB DESCRIPTION:
{job_description}

Write a compelling 3-paragraph cover letter that:
1. Opens with enthusiasm and clearly states why the candidate is a great fit
2. Connects the candidate's experience to the job requirements
3. Demonstrates understanding of the organization's mission
4. Is persuasive, authentic, and avoids generic phrases

Return ONLY the cover letter text.
"""

def call_ai(cv_text, job_description):
    """Call DeepSeek API with improved prompts"""
    try:
        print("📤 Sending request to DeepSeek API...")
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(cv_text, job_description)},
            ],
            temperature=0.2,
            response_format={"type": "json_object"},
            timeout=45.0
        )
        
        raw_response = response.choices[0].message.content
        print(f"📥 Response received: {len(raw_response)} chars")
        
        result = json.loads(raw_response)
        
        if 'summary' not in result:
            result['summary'] = "Summary not provided"
        if 'experience' not in result:
            result['experience'] = {}
        if 'technical_skills' not in result:
            result['technical_skills'] = []
        if 'core_competencies' not in result:
            result['core_competencies'] = []
        
        return result
        
    except Exception as e:
        print(f"API error: {str(e)}")
        return {
            'summary': 'Error. Please try again.',
            'experience': {},
            'technical_skills': [],
            'core_competencies': []
        }

def tailor_cover_letter(cv_text, job_description):
    """Generate tailored cover letter"""
    try:
        print("✍️ Generating cover letter...")
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert cover letter writer. Every sentence should connect the candidate's experience to the job requirements."},
                {"role": "user", "content": build_cover_letter_prompt(cv_text, job_description)}
            ],
            temperature=0.3,
            timeout=30.0
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Cover letter error: {str(e)}")
        return "Dear Hiring Manager,\n\nI am writing to express my interest in this position. With over 10 years of experience in government and international development, I am confident in my ability to contribute to your organization.\n\nSincerely,\n[Your Name]"

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

def create_tailored_docx(new_summary, new_experience, core_competencies, technical_skills):
    """Create tailored CV"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(12)
    
    # Header
    doc.add_paragraph("+252611385559/+254715244144")
    doc.add_paragraph("cos.presidency@jubalandstate.so")
    doc.add_paragraph("khalifyusuf@agileanalytica.com")
    doc.add_paragraph("")
    
    # ========== SUMMARY ==========
    summary_header = doc.add_paragraph("# Summary")
    summary_header.runs[0].font.name = 'Calibri'
    summary_header.runs[0].font.size = Pt(12)
    summary_header.runs[0].bold = True
    
    if new_summary:
        summary_para = doc.add_paragraph(new_summary)
        summary_para.runs[0].font.name = 'Calibri'
        summary_para.runs[0].font.size = Pt(12)
    doc.add_paragraph("")
    
    # ========== CORE COMPETENCIES ==========
    comp_header = doc.add_paragraph("# Core Competencies")
    comp_header.runs[0].font.name = 'Calibri'
    comp_header.runs[0].font.size = Pt(12)
    comp_header.runs[0].bold = True
    
    if core_competencies:
        for comp in core_competencies:
            comp_para = doc.add_paragraph(f"• {comp}")
            comp_para.runs[0].font.name = 'Calibri'
            comp_para.runs[0].font.size = Pt(12)
    doc.add_paragraph("")
    
    # ========== EXPERIENCE ==========
    exp_header = doc.add_paragraph("# Experience")
    exp_header.runs[0].font.name = 'Calibri'
    exp_header.runs[0].font.size = Pt(12)
    exp_header.runs[0].bold = True
    
    for job_title in JOB_TITLES:
        # Job Title
        job_para = doc.add_paragraph(job_title)
        job_para.runs[0].font.name = 'Segoe UI'
        job_para.runs[0].font.size = Pt(12)
        job_para.runs[0].bold = True
        
        # Company
        company_para = doc.add_paragraph(COMPANY_NAMES.get(job_title, ""))
        company_para.runs[0].font.name = 'Segoe UI'
        company_para.runs[0].font.size = Pt(12)
        company_para.runs[0].italic = True
        
        # Bullets
        if job_title in new_experience and new_experience[job_title]:
            for bullet in new_experience[job_title]:
                bullet_para = doc.add_paragraph(f"• {bullet}")
                bullet_para.runs[0].font.name = 'Calibri'
                bullet_para.runs[0].font.size = Pt(12)
        else:
            # Fallback bullets
            fallback_bullets = {
                "Chief of Staff (Feb 2023-To Date)": [
                    "Led strategic coordination across 15 government ministries.",
                    "Managed stakeholder relationships with international donors and NGOs.",
                    "Developed and implemented strategic plans.",
                    "Oversaw program monitoring and reporting.",
                    "Provided crisis management and emergency response coordination.",
                    "Facilitated resource mobilization and partnership building."
                ],
                "Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)": [
                    "Led humanitarian needs assessments using research methodologies.",
                    "Managed donor-funded programs with strict compliance.",
                    "Conducted analysis and reporting for policy decisions.",
                    "Established partnerships with UN agencies and NGOs.",
                    "Strengthened institutional capacity through training.",
                    "Developed monitoring frameworks and reporting systems."
                ],
                "Chief Operations Officer (Jul 2016 – Jul 2021)": [
                    "Directed operational functions including budgeting and planning.",
                    "Led business development and market expansion.",
                    "Established strategic partnerships.",
                    "Managed stakeholder relations.",
                    "Oversaw compliance and reporting.",
                    "Built and led teams."
                ]
            }
            for bullet in fallback_bullets.get(job_title, []):
                bullet_para = doc.add_paragraph(f"• {bullet}")
                bullet_para.runs[0].font.name = 'Calibri'
                bullet_para.runs[0].font.size = Pt(12)
        
        doc.add_paragraph("")
    
    # ========== TECHNICAL SKILLS ==========
    tech_header = doc.add_paragraph("# Technical Skills")
    tech_header.runs[0].font.name = 'Calibri'
    tech_header.runs[0].font.size = Pt(12)
    tech_header.runs[0].bold = True
    
    if technical_skills:
        for skill in technical_skills:
            skill_para = doc.add_paragraph(f"• {skill}")
            skill_para.runs[0].font.name = 'Calibri'
            skill_para.runs[0].font.size = Pt(12)
    doc.add_paragraph("")
    
    # ========== EDUCATION ==========
    edu_header = doc.add_paragraph("# Education")
    edu_header.runs[0].font.name = 'Calibri'
    edu_header.runs[0].font.size = Pt(12)
    edu_header.runs[0].bold = True
    
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
    
    # ========== REFEREES ==========
    ref_header = doc.add_paragraph("# REFREES")
    ref_header.runs[0].font.name = 'Calibri'
    ref_header.runs[0].font.size = Pt(12)
    ref_header.runs[0].bold = True
    
    referees = [
        "Abshir olow", "Former Chief of Staff", "Jubaland State, Somalia",
        "+254722877178", "abshirmabdi@gmail.com", "",
        "Mr. Abdirahman Abdi", "Planning Minister, Jubaland",
        "+252612992848", "mopic@jubalandstate.so", "",
        "Omar Abdi Mohamed, HDmls, Msc.Int.",
        "Commodity Logistics Director-USAID", "omarabdi2@gmail.com"
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
    
    cv_text = MASTER_CV_TEMPLATE
    
    print(f"📄 CV text length: {len(cv_text)}")
    
    try:
        print("🤖 Tailoring CV...")
        result = call_ai(cv_text, job_description)
        
        tailored_summary = result.get('summary', '')
        tailored_experience = result.get('experience', {})
        core_competencies = result.get('core_competencies', [])
        technical_skills = result.get('technical_skills', [])
        
        print(f"📝 Summary: {tailored_summary[:100] if tailored_summary else 'EMPTY'}...")
        print(f"📝 Experience keys: {list(tailored_experience.keys())}")
        
    except Exception as e:
        print(f"❌ AI error: {str(e)}")
        return None, None, f"AI error: {str(e)}", "Error", {}, [], []
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter(cv_text, job_description)
        print(f"✅ Cover letter: {len(tailored_cover)} chars")
    except Exception as e:
        tailored_cover = "Dear Hiring Manager,\n\nI am writing to express my interest in this position. With over 10 years of experience in government and international development, I am confident in my ability to contribute to your organization.\n\nSincerely,\n[Your Name]"
        print(f"⚠️ Cover letter error: {str(e)}")
    
    try:
        print("📄 Creating DOCX...")
        cv_output = create_tailored_docx(tailored_summary, tailored_experience, core_competencies, technical_skills)
        
        cover_doc = Document()
        for line in tailored_cover.split('\n'):
            if line.strip():
                cover_doc.add_paragraph(line)
        cover_output = BytesIO()
        cover_doc.save(cover_output)
        cover_output.seek(0)
        
        print("✅ Documents generated")
        return cv_output, cover_output, "Success", tailored_summary, tailored_experience, tailored_cover
        
    except Exception as e:
        print(f"❌ Document error: {str(e)}")
        traceback.print_exc()
        return None, None, f"Document error: {str(e)}", "Error", {}, "Error"

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
        
        print(f"📝 Processing job: {len(job_description)} chars")
        cv_output, cover_output, message, tailored_summary, tailored_experience, tailored_cover = process_application(job_description)
        
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
                'cover_data': cover_base64,
                'summary': tailored_summary,
                'experience': tailored_experience,
                'cover_letter': tailored_cover
            })
        else:
            return jsonify({'error': message}), 500
            
    except Exception as e:
        print(f"❌ Error: {str(e)}")
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
