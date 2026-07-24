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
    """Extract text and preserve structure"""
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

def get_docx_structure(file_path):
    """Get full document with paragraph indices for debugging"""
    doc = Document(file_path)
    result = []
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            result.append((i, para.text))
    return result, doc

def tailor_cv_deep(cv_text, job_description):
    """Rewrite Summary, Skills, and Experience bullets"""
    
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
2. Rewrite the skills section as a comma-separated list.
3. For EACH job, rewrite the bullet points to highlight relevant achievements.
4. Do NOT invent new jobs, achievements, or numbers.
5. Keep job titles, employers, and dates exactly as they appear.
6. Use the EXACT job title from the Master CV as keys.

Return JSON:
{{
    "tailored_summary": "new summary here",
    "tailored_skills": "skill 1, skill 2, skill 3, ...",
    "tailored_experience": {{
        "Chief of Staff (Feb 2023-To Date)": [
            "bullet 1",
            "bullet 2",
            "bullet 3"
        ],
        "Senior Advisor -- Projects Planning & Grants Development | Aug 2021 -- Jan 2023": [
            "bullet 1",
            "bullet 2"
        ]
    }}
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
    Update only summary, skills, and experience bullet points.
    Preserves job titles and formatting.
    """
    doc = Document(template_path)
    
    # Find all section positions
    summary_pos = -1
    skills_pos = -1
    experience_pos = -1
    education_pos = -1
    
    print("\n🔍 Scanning document headers:")
    for i, para in enumerate(doc.paragraphs):
        text = para.text.lower().strip()
        if 'summary' in text and len(text) < 30:
            summary_pos = i
            print(f"  Summary at {i}: '{para.text[:30]}'")
        elif 'skill' in text and len(text) < 30:
            skills_pos = i
            print(f"  Skills at {i}: '{para.text[:30]}'")
        elif 'experience' in text and len(text) < 30:
            experience_pos = i
            print(f"  Experience at {i}: '{para.text[:30]}'")
        elif 'education' in text and len(text) < 30:
            education_pos = i
            print(f"  Education at {i}: '{para.text[:30]}'")
    
    # 1. UPDATE SUMMARY
    if summary_pos != -1 and new_summary:
        print(f"\n📝 Updating summary at position {summary_pos}")
        end_pos = skills_pos if skills_pos > summary_pos else len(doc.paragraphs)
        
        # Clear existing summary content (keep header)
        for i in range(summary_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Insert new summary
        if summary_pos + 1 < len(doc.paragraphs):
            doc.paragraphs[summary_pos + 1].text = new_summary
            print(f"✅ Summary updated")
    
    # 2. UPDATE SKILLS - Complete rebuild
    if skills_pos != -1 and new_skills:
        print(f"\n📝 Updating skills at position {skills_pos}")
        end_pos = experience_pos if experience_pos > skills_pos else len(doc.paragraphs)
        
        # Remove entire skills section content (keep header)
        for i in range(skills_pos + 1, end_pos):
            if i < len(doc.paragraphs):
                doc.paragraphs[i].text = ""
        
        # Insert new skills as bullet points
        skills_list = [s.strip() for s in new_skills.split(',') if s.strip()]
        for i, skill in enumerate(skills_list):
            if skills_pos + 1 + i < len(doc.paragraphs):
                doc.paragraphs[skills_pos + 1 + i].text = f"• {skill}"
        print(f"✅ Skills updated with {len(skills_list)} skills")
    
    # 3. UPDATE EXPERIENCE BULLETS
    if experience_pos != -1 and new_experience:
        print(f"\n📝 Updating experience bullets")
        
        # Find all job titles in the document
        job_titles = []
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            # Detect job titles (bold with dates in parentheses)
            if '(' in text and ')' in text and any(y in text for y in ['2023', '2022', '2021', '2020', '2019', '2018', '2017', '2016']):
                # Check if it's a job title (not just any text with a date)
                if any(keyword in text.lower() for keyword in ['chief', 'senior', 'advisor', 'officer', 'manager', 'director', 'coordinator']):
                    job_titles.append((i, text))
                    print(f"  Found job: '{text[:50]}...'")
        
        # Process each job
        for pos, title in job_titles:
            # Find matching key in new_experience
            matched_key = None
            for key in new_experience.keys():
                # Check if title contains key or key contains title (fuzzy match)
                if key.lower() in title.lower() or title.lower() in key.lower():
                    matched_key = key
                    break
            
            if not matched_key:
                print(f"  ⚠️ No match for: '{title[:40]}'")
                continue
            
            new_bullets = new_experience[matched_key]
            if not new_bullets:
                continue
            
            print(f"  ✅ Updating: '{matched_key[:40]}...' with {len(new_bullets)} bullets")
            
            # Find the end of this job section
            end_pos = len(doc.paragraphs)
            for next_pos, _ in job_titles:
                if next_pos > pos:
                    end_pos = next_pos
                    break
            
            # Find bullet points in this section
            bullet_indices = []
            for i in range(pos + 1, end_pos):
                text = doc.paragraphs[i].text.strip()
                if text.startswith('-') or text.startswith('•'):
                    bullet_indices.append(i)
            
            # Replace existing bullets
            for i, bullet in enumerate(new_bullets):
                if i < len(bullet_indices):
                    # Replace existing bullet
                    idx = bullet_indices[i]
                    doc.paragraphs[idx].text = f"• {bullet}"
                else:
                    # Add new bullet at the end of the section
                    insert_pos = bullet_indices[-1] + 1 if bullet_indices else pos + 1
                    if insert_pos < len(doc.paragraphs):
                        doc.paragraphs[insert_pos].text = f"• {bullet}"
                        bullet_indices.append(insert_pos)
            
            print(f"     Inserted {len(new_bullets)} bullets")
    
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
