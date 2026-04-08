from flask import Flask, render_template, request, jsonify
import os
import re

import docx
import pdfplumber

app = Flask(__name__)

# -------------------- JOB DATABASE --------------------
JOBS = {
    "Frontend Developer": {
        "icon": "🎨",
        "skills": ["html", "css", "javascript", "react", "responsive design", "tailwind", "bootstrap", "ui", "git"]
    },
    "Backend Developer": {
        "icon": "⚙️",
        "skills": ["python", "java", "node.js", "flask", "django", "api", "sql", "mongodb", "git"]
    },
    "Full Stack Developer": {
        "icon": "🧩",
        "skills": ["html", "css", "javascript", "react", "node.js", "python", "flask", "sql", "git"]
    },
    "Data Scientist": {
        "icon": "📊",
        "skills": ["python", "pandas", "numpy", "machine learning", "sql", "statistics", "matplotlib", "data analysis"]
    },
    "Machine Learning Engineer": {
        "icon": "🤖",
        "skills": ["python", "machine learning", "tensorflow", "pytorch", "numpy", "pandas", "deep learning", "sql"]
    },
    "AI Engineer": {
        "icon": "🧠",
        "skills": ["python", "machine learning", "deep learning", "llm", "nlp", "tensorflow", "pytorch", "api"]
    },
    "Python Developer": {
        "icon": "🐍",
        "skills": ["python", "flask", "django", "oop", "sql", "api", "git", "debugging"]
    },
    "Java Developer": {
        "icon": "☕",
        "skills": ["java", "oop", "spring", "hibernate", "sql", "api", "git"]
    },
    "C++ Developer": {
        "icon": "💻",
        "skills": ["c++", "dsa", "oop", "algorithms", "debugging", "git"]
    },
    "Data Analyst": {
        "icon": "📈",
        "skills": ["python", "sql", "excel", "power bi", "tableau", "statistics", "pandas", "numpy", "data visualization"]
    },
    "DevOps Engineer": {
        "icon": "🚀",
        "skills": ["linux", "docker", "kubernetes", "terraform", "aws", "azure", "jenkins", "ci/cd", "bash"]
    },
    "Cloud Engineer": {
        "icon": "☁️",
        "skills": ["aws", "azure", "gcp", "docker", "kubernetes", "terraform", "linux", "networking", "security"]
    },
    "Cyber Security Analyst": {
        "icon": "🔐",
        "skills": ["network security", "ethical hacking", "owasp", "siem", "incident response", "cryptography", "linux"]
    },
    "UI/UX Designer": {
        "icon": "🖌️",
        "skills": ["figma", "ui/ux", "wireframing", "prototyping", "design thinking", "communication", "collaboration"]
    },
    "QA Engineer": {
        "icon": "🧪",
        "skills": ["unit testing", "manual testing", "automation testing", "selenium", "jira", "agile", "bug tracking"]
    },
    "Data Engineer": {
        "icon": "🏗️",
        "skills": ["python", "sql", "spark", "hadoop", "airflow", "kafka", "dbt", "aws", "etl"]
    },
    "Blockchain Developer": {
        "icon": "⛓️",
        "skills": ["solidity", "javascript", "node.js", "smart contracts", "web3", "ethereum", "git"]
    },
    "Game Developer": {
        "icon": "🎮",
        "skills": ["c++", "c#", "unity", "unreal engine", "oop", "algorithms", "physics"]
    },
    "Android Developer": {
        "icon": "📱",
        "skills": ["java", "kotlin", "android", "firebase", "sqlite", "rest api", "git", "oop"]
    },
    "iOS Developer": {
        "icon": "🍎",
        "skills": ["swift", "ios", "xcode", "rest api", "git", "oop"]
    },
}

# -------------------- TEXT EXTRACTION --------------------
def extract_text_from_txt(file):
    return file.read().decode("utf-8", errors="ignore")

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print("DOCX extraction error:", e)
        return ""

def extract_text_from_pdf(file):
    try:
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + " "
        return re.sub(r'\s+', ' ', text).strip()
    except Exception as e:
        print("PDF extraction error:", e)
        return ""

def extract_resume_text(uploaded_file):
    if not uploaded_file:
        return ""
    filename = uploaded_file.filename.lower()
    if filename.endswith(".txt"):
        return extract_text_from_txt(uploaded_file)
    elif filename.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)
    elif filename.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)
    else:
        return ""

# -------------------- SKILL MATCHING --------------------
def clean_text(text):
    return re.sub(r'[^a-zA-Z0-9+#.\-\s]', ' ', text.lower())

def extract_skills_from_resume(resume_text, all_skills):
    resume_text = clean_text(resume_text)
    found = set()
    for skill in all_skills:
        skill_pattern = re.escape(skill.lower())
        if re.search(r'\b' + skill_pattern + r'\b', resume_text):
            found.add(skill.lower())
    return found

def generate_ai_feedback(score, matching, missing, extra):
    strengths = []
    suggestions = []
    recommendations = []

    if len(matching) >= 5:
        strengths.append("Your resume already reflects a strong technical foundation for this role.")
    if extra:
        strengths.append("You have additional transferable skills that make your profile more versatile.")
    if "project" in " ".join(extra + matching):
        strengths.append("Project-based experience can help demonstrate practical capability.")
    if not strengths:
        strengths.append("Resume submitted successfully — review the skill gaps below to improve your match.")

    if missing:
        suggestions.append(f"Focus on learning: {', '.join(missing[:5])}.")
    if score < 70:
        suggestions.append("Add stronger role-specific keywords to improve alignment.")
    suggestions.append("Quantify project outcomes and responsibilities wherever possible.")

    if missing:
        recommendations.append(f"Build 1–2 projects using {', '.join(missing[:3])}.")
    recommendations.append("Tailor your resume summary to the selected job role.")
    recommendations.append("Highlight tools, frameworks, and measurable results more clearly.")

    return strengths, suggestions, recommendations

# -------------------- ROUTES --------------------
@app.route("/")
def home():
    return render_template("index.html", jobs=JOBS)

@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        job = request.form.get("job")
        resume_text = request.form.get("resumeText", "").strip()
        resume_file = request.files.get("resumeFile")

        if not job or job not in JOBS:
            return jsonify({"error": "Invalid or missing job role."}), 400

        if not resume_text and resume_file:
            resume_text = extract_resume_text(resume_file)

        print("RESUME LENGTH:", len(resume_text))
        print("RESUME PREVIEW:", resume_text[:500])

        if not resume_text.strip():
            return jsonify({"error": "Could not extract resume text. Try using pasted text or a clearer file."}), 400

        required_skills = [s.lower() for s in JOBS[job]["skills"]]

        all_possible_skills = set()
        for role in JOBS.values():
            all_possible_skills.update([s.lower() for s in role["skills"]])

        found_skills = extract_skills_from_resume(resume_text, all_possible_skills)

        print("ALL POSSIBLE SKILLS:", sorted(list(all_possible_skills)))
        print("FOUND SKILLS:", sorted(list(found_skills)))

        matching_skills = sorted(list(set(required_skills) & found_skills))
        missing_skills = sorted(list(set(required_skills) - found_skills))
        extra_skills = sorted(list(found_skills - set(required_skills)))

        system_score = round((len(matching_skills) / max(len(required_skills), 1)) * 100)
        ai_score = min(100, max(35, system_score + (5 if len(extra_skills) >= 3 else 0)))
        final_score = round((system_score * 0.7) + (ai_score * 0.3))

        strengths, suggestions, recommendations = generate_ai_feedback(
            final_score, matching_skills, missing_skills, extra_skills
        )

        return jsonify({
            "job": job,
            "required_skills": required_skills,
            "matching_skills": matching_skills,
            "missing_skills": missing_skills,
            "extra_skills": extra_skills,
            "system_score": system_score,
            "ai_score": ai_score,
            "final_score": final_score,
            "strengths": strengths,
            "suggestions": suggestions,
            "recommendations": recommendations
        })

    except Exception as e:
        print("ANALYZE ROUTE ERROR:", e)
        return jsonify({"error": f"Server error: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))