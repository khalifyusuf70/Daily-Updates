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
    base_url="https://api.deepseek.com/v1"
)

COVER_PATH = "Cover_Template.docx"

# ---------------------------
# MASTER CV TEMPLATE - WITH EXACT JOB TITLES
# ---------------------------
MASTER_CV_TEMPLATE = """
+252611385559/+254715244144

cos.presidency@jubalandstate.so

khalifyusuf@agileanalytica.com

# Summary
I am a highly efficient professional with over 10 years of senior leadership experience managing complex portfolios across government and international development sectors. Proven ability to lead humanitarian response, civil society strengthening, and peacebuilding programs across the Horn of Africa. Expert in strategic leadership, programme growth, resource mobilization, and donor engagement. Strong track record in partnership development, locally led programming, and organizational transformation. Skilled in political analysis, contextual understanding, and managing multi-country operations.

# Skill Highlights
[Skills will be inserted here]

# Experience

Chief of Staff (Feb 2023-To Date)
Jubaland State of Somalia
- Donor Engagement & Fundraising: Led resource mobilization efforts, engaging with international donors (USAID, EU, UN), diplomatic missions, and development partners to secure funding for state priorities and emergency response.
- Strategic Partnership Development: Established and maintained partnerships with international NGOs, UN agencies, and civil society organizations to enhance governance, humanitarian response, and service delivery.
- Strategic Leadership: Directed the Office of the President, coordinating activities across 15 government ministries and agencies to ensure alignment with the State Development Plan.
- Program Growth & Strategy: Led the development and implementation of the comprehensive State Development Plan, translating high-level goals into actionable programs and policies.
- Stakeholder Management: Represented the President at national and regional forums, managing complex relationships and ensuring aligned engagement across all partners.
- Crisis Management: Spearheaded emergency response efforts during political instability and humanitarian crises, coordinating relief activities and resource allocation.
- Established and maintained partnerships with civil society organizations and community leaders to enhance governance and service delivery.

Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)
Ministry of Planning, Jubaland State, Somalia
- Grant Acquisition & Management: Led the development of the Ministry's Strategic Plan and managed a portfolio of donor-funded programs, ensuring effective implementation, financial accountability, and compliance.
- Proposal Development: Led humanitarian and development needs assessments, translating findings into priority program areas and successful funding proposals for international donors.
- Strategy & Policy: Provided strategic leadership, advising the Minister on planning, policy, and program development to align Ministry initiatives with national and regional priorities.
- External Relations: Established and maintained partnerships with international development organizations, UN agencies, and NGOs, fostering collaboration on civil society strengthening and peacebuilding initiatives.
- Analytical Reporting: Conducted political, economic, and social analysis to inform Ministry strategy, providing critical data-driven insights for programmatic decision-making.

Chief Operations Officer (Jul 2016 – Jul 2021)
KIMS MICROFINANCE, Somalia
- Organizational Leadership: Directed all operational functions, including budgeting, strategic planning, and resource allocation, to drive organizational development and growth.
- Business Development: Led business development initiatives, expanding the organization's reach into new regions and securing funding for microenterprise programming that supported vulnerable populations.
- Partnership Management: Established and maintained strategic partnerships with international development organizations to enhance program impact and sustainability.
- Community Engagement: Managed community relations, contributing to peacebuilding and social development through economic empowerment and locally-led programming.
- Stakeholder Relations: Represented the organization to government agencies, donors, and partners, ensuring effective communication and stakeholder management.

# Education
Micro Masters. Database management System
University of Baltimore Maryland
Degree Actuarial Science and Statistics
JOMO KENYATTA UNIVERSITY
Nairobi, Kenya

# REFREES
Abshir olow
Former Chief of Staff
Jubaland State, Somalia
+254722877178
abshirmabdi@gmail.com

Mr. Abdirahman Abdi
Planning Minister, Jubaland
+252612992848
mopic@jubalandstate.so

Omar Abdi Mohamed, HDmls, Msc.Int.
Commodity Logistics Director-USAID
omarabdi2@gmail.com
"""

# ---------------------------
# EXACT JOB TITLES FOR AI PROMPT
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
    # Build experience template with exact job titles
    exp_template = {}
    for title in JOB_TITLES:
        exp_template[title] = ["bullet 1", "bullet 2", "bullet 3"]
    
    titles_json = json.dumps(JOB_TITLES, indent=2)
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
    "summary": "your new summary here",
    "experience": {exp_json}
}}
"""

def call_ai(cv_text, job_description):
    """Call DeepSeek API with the tailored prompts"""
    try:
        print("📤 Sending request to DeepSeek API...")
        response = client.chat.completions.create(
            model="deepseek-v4-pro",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(cv_text, job_description)},
            ],
            temperature=0.2,
            response_format={"type": "json_object"},
            timeout=60.0
        )
        
        raw_response = response.choices[0].message.content
        print(f"📥 Raw response length: {len(raw_response)}")
        print(f"📥 Raw response preview: {raw_response[:500]}...")
        
        result = json.loads(raw_response)
        print(f"✅ Parsed JSON keys: {result.keys()}")
        
        if 'summary' not in result:
            print("⚠️ 'summary' key missing in response")
            result['summary'] = "Summary not provided by AI"
        
        if 'experience' not in result:
            print("⚠️ 'experience' key missing in response")
            result['experience'] = {}
        
        return result
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON parse error: {str(e)}")
        print(f"Raw response: {raw_response}")
        return {
            'summary': 'Error parsing AI response. Please try again.',
            'experience': {}
        }
    except Exception as e:
        print(f"DeepSeek API error: {str(e)}")
        traceback.print_exc()
        return {
            'summary': f'Error: {str(e)}',
            'experience': {}
        }

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
            temperature=0.3,
            timeout=60.0
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Cover letter error: {str(e)}")
        return "Dear Hiring Manager,\n\nI am writing to express my strong interest in this position. With over 10 years of experience in government and international development, I am confident in my ability to contribute to your organization.\n\nSincerely,\n[Your Name]"

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
    """Create tailored CV from scratch with formatting"""
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(12)
    
    # Header
    doc.add_paragraph("+252611385559/+254715244144")
    doc.add_paragraph("cos.presidency@jubalandstate.so")
    doc.add_paragraph("khalifyusuf@agileanalytica.com")
    doc.add_paragraph("")
    
    # Summary
    summary_header = doc.add_paragraph("# Summary")
    summary_header.runs[0].font.name = 'Calibri'
    summary_header.runs[0].font.size = Pt(12)
    summary_header.runs[0].bold = True
    
    if new_summary:
        summary_para = doc.add_paragraph(new_summary)
        summary_para.runs[0].font.name = 'Calibri'
        summary_para.runs[0].font.size = Pt(12)
    else:
        doc.add_paragraph("Summary not available")
    doc.add_paragraph("")
    
    # Skills
    skills_header = doc.add_paragraph("# Skill Highlights")
    skills_header.runs[0].font.name = 'Calibri'
    skills_header.runs[0].font.size = Pt(12)
    skills_header.runs[0].bold = True
    
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
        for run in left_cell.paragraphs[0].runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(12)
        
        right_cell = table.cell(row_idx, 1)
        right_cell.text = skills_right[row_idx] if row_idx < len(skills_right) else ""
        for run in right_cell.paragraphs[0].runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(12)
    
    doc.add_paragraph("")
    
    # Experience - Using exact job titles
    exp_header = doc.add_paragraph("# Experience")
    exp_header.runs[0].font.name = 'Calibri'
    exp_header.runs[0].font.size = Pt(12)
    exp_header.runs[0].bold = True
    
    for job_title in JOB_TITLES:
        # Job Title - Segoe UI 12 Bold
        job_para = doc.add_paragraph(job_title)
        job_para.runs[0].font.name = 'Segoe UI'
        job_para.runs[0].font.size = Pt(12)
        job_para.runs[0].bold = True
        
        # Company - Segoe UI 12 Italic
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
            # Fallback bullets from master
            fallback_bullets = {
                "Chief of Staff (Feb 2023-To Date)": [
                    "Donor Engagement & Fundraising: Led resource mobilization efforts, engaging with international donors (USAID, EU, UN), diplomatic missions, and development partners.",
                    "Strategic Partnership Development: Established and maintained partnerships with international NGOs, UN agencies, and civil society organizations.",
                    "Strategic Leadership: Directed the Office of the President, coordinating activities across 15 government ministries and agencies."
                ],
                "Senior Advisor – Projects Planning & Grants Development (Aug 2021 – Jan 2023)": [
                    "Grant Acquisition & Management: Led the development of the Ministry's Strategic Plan and managed a portfolio of donor-funded programs.",
                    "Proposal Development: Led humanitarian and development needs assessments, translating findings into priority program areas.",
                    "Strategy & Policy: Provided strategic leadership, advising the Minister on planning, policy, and program development."
                ],
                "Chief Operations Officer (Jul 2016 – Jul 2021)": [
                    "Organizational Leadership: Directed all operational functions, including budgeting, strategic planning, and resource allocation.",
                    "Business Development: Led business development initiatives, expanding the organization's reach into new regions.",
                    "Partnership Management: Established and maintained strategic partnerships with international development organizations."
                ]
            }
            for bullet in fallback_bullets.get(job_title, []):
                bullet_para = doc.add_paragraph(f"• {bullet}")
                bullet_para.runs[0].font.name = 'Calibri'
                bullet_para.runs[0].font.size = Pt(12)
        
        doc.add_paragraph("")
    
    # Education
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
    
    # Referees
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
    
    cv_text = MASTER_CV_TEMPLATE
    
    if os.path.exists(COVER_PATH):
        cover_text = read_docx(COVER_PATH)
    else:
        cover_text = "Dear Hiring Manager,\n\n[Your cover letter will go here]\n\nSincerely,\n[Your Name]"
    
    print(f"📄 CV text length: {len(cv_text)}")
    
    try:
        print("🤖 Tailoring CV with AI...")
        result = call_ai(cv_text, job_description)
        
        tailored_summary = result.get('summary', '')
        tailored_experience = result.get('experience', {})
        
        print(f"📝 Summary: {tailored_summary[:100] if tailored_summary else 'EMPTY'}...")
        print(f"📝 Experience keys: {list(tailored_experience.keys())}")
        
    except Exception as e:
        print(f"❌ AI error: {str(e)}")
        return None, None, f"AI error: {str(e)}", "Error", {}, "Error"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter(cover_text, cv_text, job_description)
        print(f"✅ Cover letter: {len(tailored_cover)} chars")
    except Exception as e:
        tailored_cover = "Error generating cover letter. Please try again."
        print(f"⚠️ Cover letter error: {str(e)}")
    
    try:
        print("📄 Creating DOCX...")
        cv_output = create_tailored_docx(tailored_summary, tailored_experience)
        
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
