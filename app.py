from flask import Flask, render_template, request, jsonify
import os
import re
from collections import Counter

# File readers
import PyPDF2
import docx

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
    "Android Developer": {
        "icon": "📱",
        "skills": ["java", "kotlin", "android", "xml", "firebase", "api", "ui"]
    },
    "Web Developer": {
        "icon": "🌐",
        "skills": ["html", "css", "javascript", "react", "flask", "bootstrap", "git"]
    },
    "UI/UX Designer": {
        "icon": "🖌️",
        "skills": ["figma", "wireframing", "prototyping", "ui", "ux", "typography", "design system"]
    },
    "DevOps Engineer": {
        "icon": "🚀",
        "skills": ["linux", "docker", "kubernetes", "aws", "ci/cd", "bash", "git", "terraform"]
    },
    "Cloud Engineer": {
        "icon": "☁️",
        "skills": ["aws", "azure", "gcp", "cloud", "linux", "networking", "docker"]
    },
    "Cybersecurity Analyst": {
        "icon": "🔐",
        "skills": ["network security", "linux", "ethical hacking", "penetration testing", "wireshark", "python"]
    },
    "Software Tester": {
        "icon": "🧪",
        "skills": ["manual testing", "selenium", "bug tracking", "test cases", "qa", "automation testing"]
    },
    "QA Engineer": {
        "icon": "✅",
        "skills": ["qa", "selenium", "automation testing", "test cases", "bug tracking", "api testing"]
    },
    "Database Administrator": {
        "icon": "🗄️",
        "skills": ["sql", "mysql", "postgresql", "database design", "backup", "optimization"]
    },
    "Business Analyst": {
        "icon": "📈",
        "skills": ["excel", "sql", "data analysis", "communication", "requirements gathering", "power bi"]
    },
    "Data Analyst": {
        "icon": "📉",
        "skills": ["excel", "sql", "python", "power bi", "tableau", "data analysis", "statistics"]
    },
    "Game Developer": {
        "icon": "🎮",
        "skills": ["unity", "c#", "game design", "c++", "physics", "debugging"]
    },
    "Embedded Systems Engineer": {
        "icon": "🔌",
        "skills": ["c", "c++", "microcontrollers", "embedded systems", "arduino", "iot"]
    },
    "IoT Developer": {
        "icon": "📡",
        "skills": ["iot", "arduino", "raspberry pi", "embedded systems", "sensors", "c", "python"]
    },
    "Blockchain Developer": {
        "icon": "⛓️",
        "skills": ["solidity", "ethereum", "web3", "smart contracts", "javascript", "blockchain"]
    },
    "AR/VR Developer": {
        "icon": "🥽",
        "skills": ["unity", "c#", "3d", "vr", "ar", "game development"]
    },
    "Product Manager": {
        "icon": "📦",
        "skills": ["product strategy", "roadmap", "communication", "market research", "agile", "analytics"]
    },
    "Project Manager": {
        "icon": "🗂️",
        "skills": ["project management", "agile", "scrum", "leadership", "communication", "planning"]
    },
    "Technical Writer": {
        "icon": "✍️",
        "skills": ["documentation", "technical writing", "api docs", "communication", "research"]
    },
    "System Administrator": {
        "icon": "🖥️",
        "skills": ["linux", "windows server", "networking", "troubleshooting", "security", "bash"]
    }
}

# -------------------- TEXT EXTRACTION --------------------
def extract_text_from_txt(file):
    return file.read().decode("utf-8", errors="ignore")

def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
        return text
    except Exception as e:
        print("PDF extraction error:", e)
        return ""

def extract_text_from_docx(file):
    try:
        doc = docx.Document(file)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print("DOCX extraction error:", e)
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
    return re.sub(r'[^a-zA-Z0-9+#.\s]', ' ', text.lower())

def extract_skills_from_resume(resume_text, all_skills):
    resume_text = clean_text(resume_text)
    found = set()

    for skill in all_skills:
        if skill.lower() in resume_text:
            found.add(skill.lower())

    return found

def generate_ai_feedback(score, matching, missing, extra):
    strengths = []
    suggestions = []
    recommendations = []

    if len(matching) >= 5:
        strengths.append("Your resume already reflects a strong technical foundation for this role.")
    if extra:
        strengths.append("You have additional transferable skills that can make your profile more versatile.")
    if "projects" in " ".join(extra + matching):
        strengths.append("Project-based experience can help demonstrate practical capability.")

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

        if not resume_text.strip():
            return jsonify({"error": "Could not extract resume text. Try using pasted text or a clearer file."}), 400

        required_skills = [skill.lower() for skill in JOBS[job]["skills"]]

        all_possible_skills = set()
        for role in JOBS.values():
            all_possible_skills.update([s.lower() for s in role["skills"]])

        found_skills = extract_skills_from_resume(resume_text, all_possible_skills)

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
    app.run(debug=True, use_reloader=False)