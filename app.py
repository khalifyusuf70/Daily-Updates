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
            model="deepseek-v4-pro",
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

def extract_job_titles(cv_text):
    """Extract exact job titles from CV text"""
    job_titles = []
    lines = cv_text.split('\n')
    for line in lines:
        # Look for job title patterns
        if '(' in line and ')' in line and any(y in line for y in ['2023', '2022', '2021', '2020', '2019', '2018', '2017', '2016']):
            if any(keyword in line.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager', 'director', 'coordinator']):
                # Clean the line
                clean = re.sub(r'\{#.*?\}', '', line)
                clean = re.sub(r'\.Styl\d+', '', clean)
                clean = re.sub(r'#', '', clean)
                if clean.strip():
                    job_titles.append(clean.strip())
    return job_titles

def tailor_cv_deep(cv_text, job_description):
    """Rewrite Summary, Skills, and Experience bullets"""
    
    # Extract exact job titles from CV
    job_titles = extract_job_titles(cv_text)
    print(f"📌 Extracted job titles: {job_titles}")
    
    # Build the experience template with exact job titles
    exp_template = {}
    for title in job_titles:
        exp_template[title] = ["bullet 1", "bullet 2"]
    
    # Generate skills list for template
    skills_template = "skill 1, skill 2, skill 3, skill 4, skill 5, skill 6, skill 7, skill 8, skill 9, skill 10"
    
    prompt = f"""
You are a professional CV tailoring expert.

MASTER CV:
{cv_text}

JOB DESCRIPTION:
{job_description}

TASK:
Rewrite the CV to align with the job description, while being 100% truthful.

RULES:
1. Rewrite the professional summary (4-6 sentences).
2. Rewrite the skills section as a comma-separated list of 10-12 skills.
3. For EACH job below, rewrite the bullet points to highlight relevant achievements.
4. Do NOT invent new jobs, achievements, or numbers.
5. Keep job titles, employers, and dates exactly as they appear.

CRITICAL: Use ONLY these exact job titles as keys (copy them exactly):
{json.dumps(job_titles, indent=2)}

Return JSON:
{{
    "tailored_summary": "new summary here",
    "tailored_skills": "{skills_template}",
    "tailored_experience": {json.dumps(exp_template, indent=2)}
}}
"""
    return call_deepseek(prompt)

def tailor_cover_letter_deep(cover_text, cv_text, job_description):
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
            temperature=0.4
        )
        return response.choices[0].message.content
    except Exception as e:
        print(f"Cover letter error: {str(e)}")
        return cover_text

def update_docx_sections(template_path, new_summary, new_skills, new_experience):
    """
    Update summary, skills, and experience using direct text replacement
    """
    doc = Document(template_path)
    
    # Find section positions
    summary_pos = -1
    skills_pos = -1
    experience_pos = -1
    
    print("\n🔍 Finding sections:")
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'summary' in text and len(text) < 30:
            summary_pos = i
            print(f"  Summary at {i}")
        elif 'skill' in text and len(text) < 30:
            skills_pos = i
            print(f"  Skills at {i}")
        elif 'experience' in text and len(text) < 30:
            experience_pos = i
            print(f"  Experience at {i}")
            break
    
    # 1. UPDATE SUMMARY
    if summary_pos != -1 and new_summary:
        print(f"\n📝 Updating summary...")
        end_pos = skills_pos if skills_pos > summary_pos else len(doc.paragraphs)
        
        # Clear existing content
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Insert new summary
        if summary_pos + 1 < len(doc.paragraphs):
            doc.paragraphs[summary_pos + 1].text = new_summary
            print(f"✅ Summary updated")
    
    # 2. UPDATE SKILLS
    if skills_pos != -1 and new_skills:
        print(f"\n📝 Updating skills...")
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        
        # Clear existing content
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Insert new skills as bullet points
        skills_list = [s.strip() for s in new_skills.split(',') if s.strip()]
        for i, skill in enumerate(skills_list):
            if skills_pos + 1 + i < len(doc.paragraphs):
                doc.paragraphs[skills_pos + 1 + i].text = f"• {skill}"
        print(f"✅ Skills updated with {len(skills_list)} skills")
    
    # 3. UPDATE EXPERIENCE - Improved with fuzzy matching
    if experience_pos != -1 and new_experience:
        print(f"\n📝 Updating experience...")
        
        # Create a list of all paragraph texts with their indices
        paragraphs = [(i, para.text.strip()) for i, para in enumerate(doc.paragraphs)]
        
        # For each job from AI, try to find it in the document
        for job_title, new_bullets in new_experience.items():
            print(f"  Looking for: {job_title[:40]}...")
            
            # Try multiple matching strategies
            job_idx = -1
            
            # Strategy 1: Exact match
            for i, text in paragraphs:
                if text == job_title:
                    job_idx = i
                    break
            
            # Strategy 2: Contains match (AI key is in document text)
            if job_idx == -1:
                for i, text in paragraphs:
                    # Clean both strings for comparison
                    clean_title = re.sub(r'[^a-zA-Z0-9 ]', '', job_title.lower())
                    clean_text = re.sub(r'[^a-zA-Z0-9 ]', '', text.lower())
                    if clean_title in clean_text or clean_text in clean_title:
                        job_idx = i
                        break
            
            # Strategy 3: Keyword matching (Chief of Staff, Senior Advisor, etc.)
            if job_idx == -1:
                # Extract keywords from job title
                keywords = re.findall(r'\b(Chief|Senior|Advisor|Officer|Manager|Director|Coordinator)\b', job_title, re.IGNORECASE)
                for i, text in paragraphs:
                    if any(keyword.lower() in text.lower() for keyword in keywords):
                        # Also check if year ranges match
                        if any(str(year) in text for year in range(2016, 2024)):
                            job_idx = i
                            break
            
            # Strategy 4: Partial word matching (70% of words match)
            if job_idx == -1:
                title_words = set(re.findall(r'\b\w+\b', job_title.lower()))
                for i, text in paragraphs:
                    text_words = set(re.findall(r'\b\w+\b', text.lower()))
                    if len(title_words.intersection(text_words)) > len(title_words) * 0.6:
                        job_idx = i
                        break
            
            if job_idx == -1:
                print(f"    ⚠️ Job not found: {job_title[:40]}")
                continue
            
            print(f"    ✅ Found at index {job_idx}: {doc.paragraphs[job_idx].text[:40]}")
            
            # Find the end of this job section
            end_idx = len(doc.paragraphs)
            # Look for next job title (any of the AI keys)
            for next_job in new_experience.keys():
                if next_job == job_title:
                    continue
                for i, text in paragraphs:
                    if i > job_idx and next_job in text:
                        if i < end_idx:
                            end_idx = i
                            break
            
            # Also stop at Education or References
            for i, text in paragraphs:
                if i > job_idx and ('education' in text.lower() or 'refree' in text.lower()):
                    if i < end_idx:
                        end_idx = i
                        break
            
            print(f"    Section: {job_idx+1} to {end_idx-1}")
            
            # Find all bullet points in this section
            bullet_indices = []
            for i in range(job_idx + 1, end_idx):
                if i < len(doc.paragraphs):
                    text = doc.paragraphs[i].text.strip()
                    # Check if this is a bullet point
                    is_bullet = False
                    if text.startswith('-') or text.startswith('•') or text.startswith('*'):
                        is_bullet = True
                    elif len(text) > 5 and text[0].isdigit() and text[1] == '.':
                        is_bullet = True
                    elif text.startswith('o ') or text.startswith('> '):
                        is_bullet = True
                    # Check if it's a bullet with spaces
                    elif text.startswith('  -') or text.startswith('  •'):
                        is_bullet = True
                    
                    # Skip employer lines (company names in italics)
                    if any(keyword in text for keyword in ['Jubaland State', 'Ministry of Planning', 'KIMS MICROFINANCE', 'Jubaland']):
                        is_bullet = False
                    
                    if is_bullet:
                        bullet_indices.append(i)
            
            print(f"    Found {len(bullet_indices)} bullet points")
            
            # Replace bullets
            for i, bullet_text in enumerate(new_bullets):
                if i < len(bullet_indices):
                    idx = bullet_indices[i]
                    if idx < len(doc.paragraphs):
                        doc.paragraphs[idx].text = f"• {bullet_text}"
                        print(f"      Replaced bullet {i+1}")
                else:
                    # Add new bullet
                    insert_pos = bullet_indices[-1] + 1 if bullet_indices else job_idx + 2
                    if insert_pos < len(doc.paragraphs) and insert_pos < end_idx:
                        doc.paragraphs[insert_pos].text = f"• {bullet_text}"
                        print(f"      Added bullet {i+1}")
                    elif insert_pos < len(doc.paragraphs):
                        # Insert at position
                        new_para = doc.paragraphs[insert_pos]
                        new_para.text = f"• {bullet_text}"
                        print(f"      Added bullet {i+1}")
            
            print(f"    ✅ Updated with {len(new_bullets)} bullets")
    
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
        print("🤖 Tailoring CV with DeepSeek...")
        result = tailor_cv_deep(cv_text, job_description)
        
        tailored_summary = result.get('tailored_summary', '')
        tailored_skills = result.get('tailored_skills', '')
        tailored_experience = result.get('tailored_experience', {})
        
        print(f"📝 Summary: {tailored_summary[:100]}...")
        print(f"📝 Skills: {tailored_skills[:100]}...")
        print(f"📝 Jobs: {list(tailored_experience.keys())}")
        
    except Exception as e:
        return None, None, f"AI tailoring failed: {str(e)}"
    
    try:
        print("✍️ Generating cover letter...")
        tailored_cover = tailor_cover_letter_deep(cover_text, cv_text, job_description)
    except Exception as e:
        tailored_cover = cover_text
        print(f"⚠️ Cover letter failed: {str(e)}")
    
    try:
        print("📄 Generating tailored CV...")
        cv_output = update_docx_sections(CV_PATH, tailored_summary, tailored_skills, tailored_experience)
        
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
