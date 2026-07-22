import os
import json
import re
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
from openai import OpenAI

# Railway will set this environment variable
PORT = int(os.environ.get("PORT", 8000))

# Initialize OpenAI with Railway environment variable
API_KEY = os.environ.get("OPENAI_API_KEY")
if not API_KEY:
    raise ValueError("OPENAI_API_KEY environment variable not set")
client = OpenAI(api_key=API_KEY)

def read_docx(file_path):
    """Extract text from a .docx file"""
    doc = Document(file_path)
    paragraphs = []
    for para in doc.paragraphs:
        if para.text.strip():
            paragraphs.append(para.text)
    return "\n".join(paragraphs)

def write_docx(text, template_path, output_path):
    """Write tailored text to a new .docx preserving formatting"""
    doc = Document(template_path)
    
    # Split new text into paragraphs
    new_paragraphs = [p for p in text.split('\n') if p.strip()]
    
    # Replace paragraphs
    para_index = 0
    for paragraph in doc.paragraphs:
        if para_index >= len(new_paragraphs):
            break
        if paragraph.text.strip():
            if paragraph.runs:
                paragraph.runs[0].text = new_paragraphs[para_index]
                for run in paragraph.runs[1:]:
                    run.text = ""
                para_index += 1
    
    # Add extra paragraphs if needed
    while para_index < len(new_paragraphs):
        doc.add_paragraph(new_paragraphs[para_index])
        para_index += 1
    
    doc.save(output_path)

def tailor_cv(cv_text, job_description):
    """Send CV and job description to OpenAI for tailoring"""
    prompt = f"""
You are a professional CV tailoring assistant.

Input:
1. Master CV text
2. Job description text

Output in JSON format with one key:
- tailored_cv: The complete tailored CV text

Rules:
- NEVER add experience, skills, or qualifications that aren't in the master CV
- ONLY reword existing content to better match the job description
- Use keywords from the job description naturally
- Preserve all factual information

Master CV:
{cv_text}

Job Description:
{job_description}

Return JSON with the tailored CV:
"""
    
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
    return result.get("tailored_cv", "")

def tailor_cover_letter(cover_text, job_description, cv_text):
    """Generate a tailored cover letter"""
    prompt = f"""
You are a professional cover letter writer.

Input:
1. Cover letter template
2. Job description
3. Master CV

Output in JSON format with one key:
- tailored_cover: The complete tailored cover letter

Rules:
- Use the template structure
- Personalize with specific examples from the CV
- Match keywords from the job description
- Keep it professional and concise

Cover Letter Template:
{cover_text}

Job Description:
{job_description}

Master CV:
{cv_text}

Return JSON with the tailored cover letter:
"""
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert cover letter writer."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.3
    )
    
    result = json.loads(response.choices[0].message.content)
    return result.get("tailored_cover", "")

def process_application(cv_path, cover_path, job_description):
    """Process entire application tailoring"""
    # Read documents
    cv_text = read_docx(cv_path)
    cover_text = read_docx(cover_path)
    
    # Tailor CV
    print("📝 Tailoring CV...")
    tailored_cv = tailor_cv(cv_text, job_description)
    
    # Tailor cover letter
    print("📝 Tailoring Cover Letter...")
    tailored_cover = tailor_cover_letter(cover_text, job_description, cv_text)
    
    # Write documents
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cv_output = f"Tailored_CV_{timestamp}.docx"
    cover_output = f"Tailored_Cover_{timestamp}.docx"
    
    write_docx(tailored_cv, cv_path, cv_output)
    write_docx(tailored_cover, cover_path, cover_output)
    
    return cv_output, cover_output

def main():
    """Main entry point"""
    print("🚀 CV Tailor Service Starting...")
    
    # Check if files exist
    cv_file = "Master_CV.docx"
    cover_file = "Cover_Template.docx"
    
    if not os.path.exists(cv_file):
        print(f"⚠️ Please upload your CV as '{cv_file}'")
        return
    
    if not os.path.exists(cover_file):
        print(f"⚠️ Please upload your cover letter template as '{cover_file}'")
        return
    
    # Read job description from file or stdin
    job_desc_file = "job_description.txt"
    if os.path.exists(job_desc_file):
        with open(job_desc_file, 'r', encoding='utf-8') as f:
            job_description = f.read()
        print(f"📄 Loaded job description from {job_desc_file}")
    else:
        print("📝 Please paste the job description (type 'END' on new line to finish):")
        lines = []
        while True:
            line = input()
            if line.strip().upper() == "END":
                break
            lines.append(line)
        job_description = "\n".join(lines)
    
    if not job_description.strip():
        print("❌ No job description provided. Exiting.")
        return
    
    # Process application
    try:
        cv_output, cover_output = process_application(cv_file, cover_file, job_description)
        print(f"✅ Tailored CV saved: {cv_output}")
        print(f"✅ Tailored Cover Letter saved: {cover_output}")
        print("🎉 Processing complete!")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()
