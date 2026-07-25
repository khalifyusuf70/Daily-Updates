import os
import json
import base64
from datetime import datetime
from flask import Flask, render_template, request, jsonify
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

def tailor_cv_content(experience_text, job_description):
    """
    Generate tailored summary, experience bullets, and cover letter
    """
    
    prompt = f"""
You are a professional CV tailoring expert.

MY EXPERIENCE:
{experience_text}

JOB DESCRIPTION:
{job_description}

TASK:
Rewrite my CV content to perfectly match this job description, while being 100% truthful.

RULES:
1. Write a compelling professional summary (4-6 sentences) that matches the job.
2. For EACH job in my experience, rewrite the bullet points (4-5 per job) to highlight relevant achievements.
3. Write a personalized cover letter (3 paragraphs) that connects my experience to the job.
4. Do NOT invent new jobs, achievements, or numbers.
5. Keep job titles, employers, and dates exactly as they appear.
6. Use keywords from the job description naturally.

Return JSON:
{{
    "summary": "professional summary here",
    "experience": {{
        "Job Title 1 | Company | Date": [
            "bullet 1",
            "bullet 2",
            "bullet 3",
            "bullet 4"
        ],
        "Job Title 2 | Company | Date": [
            "bullet 1",
            "bullet 2",
            "bullet 3"
        ]
    }},
    "cover_letter": "full cover letter text here"
}}
"""
    
    return call_deepseek(prompt)

def process_application(experience_text, job_description):
    """Process the entire application"""
    
    if not experience_text or not job_description:
        return None, "Please provide both experience and job description"
    
    try:
        print("🤖 Generating tailored content...")
        result = tailor_cv_content(experience_text, job_description)
        
        summary = result.get('summary', '')
        experience = result.get('experience', {})
        cover_letter = result.get('cover_letter', '')
        
        print(f"📝 Summary: {summary[:100]}...")
        print(f"📝 Jobs: {list(experience.keys())}")
        print(f"📝 Cover letter: {len(cover_letter)} characters")
        
        return {
            'summary': summary,
            'experience': experience,
            'cover_letter': cover_letter
        }, None
        
    except Exception as e:
        return None, f"AI processing failed: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/tailor', methods=['POST'])
def tailor():
    try:
        experience_text = request.form.get('experience', '').strip()
        job_description = request.form.get('job_description', '').strip()
        
        if not experience_text or not job_description:
            return jsonify({'error': 'Please provide both your experience and the job description'}), 400
        
        print(f"📝 Processing request...")
        print(f"📝 Experience length: {len(experience_text)} characters")
        print(f"📝 Job description length: {len(job_description)} characters")
        
        result, error = process_application(experience_text, job_description)
        
        if error:
            return jsonify({'error': error}), 500
        
        return jsonify({
            'success': True,
            'message': 'Content generated successfully!',
            'summary': result['summary'],
            'experience': result['experience'],
            'cover_letter': result['cover_letter']
        })
            
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
    app.run(host='0.0.0.0', port=port, debug=False)
