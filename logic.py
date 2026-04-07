import os
import re
import json

# Optional libraries
try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    import openai
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False


# ==============================================================================
# 1. SKILLS DATABASE
# ==============================================================================

SKILLS_DATABASE = [
    "python", "java", "javascript", "typescript", "c", "c++", "c#", "r",
    "go", "rust", "swift", "kotlin", "php", "ruby", "scala", "perl",
    "matlab", "bash", "shell", "powershell", "sql", "nosql", "html", "css",

    "react", "angular", "vue", "next.js", "nuxt", "svelte", "jquery",
    "bootstrap", "tailwind", "sass", "webpack", "vite", "graphql", "rest api",

    "node.js", "express", "django", "flask", "fastapi", "spring boot",
    "asp.net", "laravel", "rails", "gin", "fiber",

    "machine learning", "deep learning", "natural language processing",
    "nlp", "computer vision", "tensorflow", "pytorch", "keras", "scikit-learn",
    "pandas", "numpy", "scipy", "matplotlib", "seaborn", "plotly",
    "hugging face", "transformers", "bert", "gpt", "llm", "reinforcement learning",
    "xgboost", "lightgbm", "catboost", "feature engineering", "statistics",

    "spark", "hadoop", "kafka", "airflow", "dbt", "dask", "hive",
    "tableau", "power bi", "looker", "excel", "google analytics",

    "mysql", "postgresql", "mongodb", "redis", "sqlite", "oracle",
    "cassandra", "dynamodb", "elasticsearch", "neo4j",

    "aws", "azure", "gcp", "docker", "kubernetes", "terraform",
    "ansible", "jenkins", "github actions", "ci/cd", "linux", "nginx",

    "git", "agile", "scrum", "jira", "unit testing", "tdd",
    "microservices", "api design", "system design", "oop",
    "data structures", "algorithms", "design patterns",

    "penetration testing", "ethical hacking", "network security",
    "siem", "vulnerability assessment", "cryptography", "firewalls",
    "owasp", "soc", "incident response",

    "communication", "leadership", "teamwork", "problem solving",
    "critical thinking", "time management", "project management",
    "presentation", "collaboration", "adaptability",

    "flutter", "react native", "firebase", "android", "ios", "xcode",
    "figma", "ui/ux", "wireframing", "prototyping", "design thinking",
    "data visualization", "mlops", "networking", "security", "database design",
    "backup", "performance tuning", "manual testing", "automation testing",
    "selenium", "bug tracking", "etl", "solidity", "smart contracts",
    "web3", "ethereum", "unity", "unreal engine", "physics",
    "microcontrollers", "arduino", "raspberry pi", "electronics", "iot",
    "mqtt", "sensors", "routing", "switching", "windows server",
    "troubleshooting"
]


# ==============================================================================
# 2. JOB ROLE → REQUIRED SKILLS + ICONS
# ==============================================================================

JOBS = {
    "Data Scientist": {
        "icon": "📊",
        "skills": [
            "python", "r", "machine learning", "deep learning", "pandas", "numpy",
            "scikit-learn", "tensorflow", "pytorch", "sql", "statistics",
            "feature engineering", "nlp", "tableau", "communication", "critical thinking"
        ]
    },
    "Data Analyst": {
        "icon": "📈",
        "skills": [
            "python", "sql", "excel", "power bi", "tableau", "statistics",
            "pandas", "numpy", "data visualization", "critical thinking", "communication"
        ]
    },
    "Machine Learning Engineer": {
        "icon": "🤖",
        "skills": [
            "python", "machine learning", "deep learning", "tensorflow", "pytorch",
            "scikit-learn", "numpy", "pandas", "mlops", "docker", "sql",
            "feature engineering", "api design", "problem solving"
        ]
    },
    "AI Engineer": {
        "icon": "🧠",
        "skills": [
            "python", "machine learning", "deep learning", "nlp", "transformers",
            "hugging face", "tensorflow", "pytorch", "llm", "gpt", "api design",
            "docker", "sql", "problem solving"
        ]
    },
    "Web Developer": {
        "icon": "🌐",
        "skills": [
            "html", "css", "javascript", "react", "typescript", "git",
            "rest api", "tailwind", "bootstrap", "problem solving", "communication"
        ]
    },
    "Frontend Developer": {
        "icon": "🎨",
        "skills": [
            "html", "css", "javascript", "typescript", "react", "next.js",
            "tailwind", "bootstrap", "sass", "webpack", "git", "problem solving"
        ]
    },
    "Backend Developer": {
        "icon": "🛠️",
        "skills": [
            "python", "java", "node.js", "express", "django", "flask", "fastapi",
            "sql", "mongodb", "postgresql", "rest api", "api design", "git",
            "problem solving"
        ]
    },
    "Full Stack Developer": {
        "icon": "🧩",
        "skills": [
            "html", "css", "javascript", "typescript", "react", "node.js",
            "express", "mongodb", "sql", "rest api", "git", "problem solving",
            "communication"
        ]
    },
    "Software Engineer": {
        "icon": "⚙️",
        "skills": [
            "python", "java", "c++", "data structures", "algorithms",
            "system design", "git", "oop", "design patterns", "sql",
            "unit testing", "agile", "microservices", "linux", "problem solving"
        ]
    },
    "Python Developer": {
        "icon": "🐍",
        "skills": [
            "python", "django", "flask", "fastapi", "sql", "postgresql",
            "api design", "oop", "git", "problem solving"
        ]
    },
    "Java Developer": {
        "icon": "☕",
        "skills": [
            "java", "spring boot", "sql", "oop", "design patterns",
            "microservices", "git", "unit testing", "problem solving"
        ]
    },
    "App Developer": {
        "icon": "📱",
        "skills": [
            "java", "kotlin", "swift", "flutter", "react native", "firebase",
            "api design", "git", "problem solving", "ui/ux"
        ]
    },
    "Android Developer": {
        "icon": "🤖",
        "skills": [
            "java", "kotlin", "android", "firebase", "sqlite", "rest api",
            "git", "problem solving", "oop"
        ]
    },
    "iOS Developer": {
        "icon": "🍎",
        "skills": [
            "swift", "ios", "xcode", "rest api", "git", "problem solving", "oop"
        ]
    },
    "DevOps Engineer": {
        "icon": "🚀",
        "skills": [
            "linux", "docker", "kubernetes", "terraform", "aws", "azure", "gcp",
            "jenkins", "github actions", "ci/cd", "bash", "ansible", "nginx"
        ]
    },
    "Cloud Engineer": {
        "icon": "☁️",
        "skills": [
            "aws", "azure", "gcp", "docker", "kubernetes", "terraform",
            "linux", "networking", "security", "problem solving"
        ]
    },
    "Cyber Security Analyst": {
        "icon": "🔐",
        "skills": [
            "network security", "ethical hacking", "owasp", "siem", "incident response",
            "cryptography", "firewalls", "vulnerability assessment", "linux",
            "problem solving"
        ]
    },
    "Penetration Tester": {
        "icon": "🕵️",
        "skills": [
            "penetration testing", "ethical hacking", "owasp", "network security",
            "linux", "python", "vulnerability assessment", "problem solving"
        ]
    },
    "Database Administrator": {
        "icon": "🗄️",
        "skills": [
            "sql", "mysql", "postgresql", "oracle", "mongodb", "database design",
            "backup", "performance tuning", "problem solving"
        ]
    },
    "Business Analyst": {
        "icon": "📋",
        "skills": [
            "excel", "sql", "tableau", "power bi", "communication",
            "critical thinking", "presentation", "problem solving", "project management"
        ]
    },
    "UI/UX Designer": {
        "icon": "🖌️",
        "skills": [
            "figma", "ui/ux", "wireframing", "prototyping", "design thinking",
            "communication", "collaboration", "problem solving"
        ]
    },
    "QA Engineer": {
        "icon": "🧪",
        "skills": [
            "unit testing", "manual testing", "automation testing", "selenium",
            "jira", "agile", "problem solving", "communication"
        ]
    },
    "Test Engineer": {
        "icon": "🧫",
        "skills": [
            "manual testing", "automation testing", "selenium", "unit testing",
            "jira", "bug tracking", "problem solving"
        ]
    },
    "Data Engineer": {
        "icon": "🏗️",
        "skills": [
            "python", "sql", "spark", "hadoop", "airflow", "kafka", "dbt",
            "aws", "azure", "gcp", "etl", "problem solving"
        ]
    },
    "Big Data Engineer": {
        "icon": "📦",
        "skills": [
            "spark", "hadoop", "kafka", "hive", "python", "sql",
            "airflow", "etl", "problem solving"
        ]
    },
    "Blockchain Developer": {
        "icon": "⛓️",
        "skills": [
            "solidity", "javascript", "node.js", "smart contracts",
            "web3", "ethereum", "git", "problem solving"
        ]
    },
    "Game Developer": {
        "icon": "🎮",
        "skills": [
            "c++", "c#", "unity", "unreal engine", "oop", "problem solving",
            "algorithms", "physics"
        ]
    },
    "Embedded Systems Engineer": {
        "icon": "🔌",
        "skills": [
            "c", "c++", "microcontrollers", "arduino", "raspberry pi",
            "electronics", "iot", "problem solving"
        ]
    },
    "IoT Developer": {
        "icon": "📡",
        "skills": [
            "c", "c++", "python", "arduino", "raspberry pi", "iot",
            "mqtt", "sensors", "problem solving"
        ]
    },
    "Network Engineer": {
        "icon": "🌍",
        "skills": [
            "networking", "linux", "firewalls", "routing", "switching",
            "security", "problem solving"
        ]
    },
    "System Administrator": {
        "icon": "🖥️",
        "skills": [
            "linux", "windows server", "networking", "bash", "powershell",
            "security", "troubleshooting", "problem solving"
        ]
    },
    "Product Manager": {
        "icon": "📦",
        "skills": [
            "communication", "leadership", "project management", "agile",
            "jira", "problem solving", "presentation", "collaboration"
        ]
    },
    "Project Manager": {
        "icon": "📁",
        "skills": [
            "project management", "leadership", "communication", "agile",
            "scrum", "jira", "time management", "collaboration"
        ]
    }
}


# ==============================================================================
# 3. FILE READING
# ==============================================================================

def read_pdf(filepath: str) -> str:
    if PyPDF2 is None:
        raise ImportError("PyPDF2 is not installed. Run: pip install PyPDF2")

    text_parts = []
    with open(filepath, "rb") as fh:
        reader = PyPDF2.PdfReader(fh)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)

    return "\n".join(text_parts)


def read_docx(filepath: str) -> str:
    if DocxDocument is None:
        raise ImportError("python-docx is not installed. Run: pip install python-docx")

    doc = DocxDocument(filepath)
    paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
    return "\n".join(paragraphs)


def read_resume(filepath: str) -> str:
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".txt":
        with open(filepath, "r", encoding="utf-8", errors="ignore") as fh:
            return fh.read()
    elif ext == ".pdf":
        return read_pdf(filepath)
    elif ext == ".docx":
        return read_docx(filepath)
    else:
        raise ValueError("Unsupported file format. Please upload .txt, .pdf, or .docx")


# ==============================================================================
# 4. SKILL EXTRACTION
# ==============================================================================

def extract_skills(resume_text: str) -> set:
    text_lower = resume_text.lower()
    found = set()

    for skill in SKILLS_DATABASE:
        if " " in skill or "." in skill or "#" in skill or "+" in skill:
            if skill in text_lower:
                found.add(skill)
        else:
            pattern = r"\b" + re.escape(skill) + r"\b"
            if re.search(pattern, text_lower):
                found.add(skill)

    return found


# ==============================================================================
# 5. SKILL GAP
# ==============================================================================

def skill_gap(extracted: set, required: list) -> dict:
    required_set = set(skill.lower() for skill in required)
    extracted_lower = set(s.lower() for s in extracted)

    matching = extracted_lower & required_set
    missing = required_set - extracted_lower
    extra = extracted_lower - required_set

    return {
        "matching": matching,
        "missing": missing,
        "extra": extra,
    }


# ==============================================================================
# 6. SCORING
# ==============================================================================

def calculate_score(gap_result: dict) -> float:
    total_required = len(gap_result["matching"]) + len(gap_result["missing"])
    if total_required == 0:
        return 0.0
    score = (len(gap_result["matching"]) / total_required) * 100
    return round(score, 2)


def final_score(system_score: float, ai_score: float,
                system_weight: float = 0.6, ai_weight: float = 0.4) -> float:
    return round(system_score * system_weight + ai_score * ai_weight, 2)


# ==============================================================================
# 7. AI FEEDBACK
# ==============================================================================

def generate_ai_feedback(resume_text: str, job_role: str, openai_api_key: str = "") -> dict:
    if HAS_OPENAI and openai_api_key:
        return _openai_feedback(resume_text, job_role, openai_api_key)
    return _mock_feedback(resume_text, job_role)


def _openai_feedback(resume_text: str, job_role: str, api_key: str) -> dict:
    openai.api_key = api_key

    prompt = f"""
You are a professional HR consultant. Evaluate the following resume for the
role of "{job_role}". Respond ONLY with a valid JSON object (no markdown):

{{
  "score": <integer 0-100>,
  "strengths": ["...", "..."],
  "weaknesses": ["...", "..."],
  "suggestions": ["...", "..."]
}}

Resume:
\"\"\"
{resume_text[:3000]}
\"\"\"
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r"```json|```", "", raw).strip()
        return json.loads(raw)
    except Exception as exc:
        mock = _mock_feedback(resume_text, job_role)
        mock["weaknesses"].append(f"(OpenAI error: {exc})")
        return mock


def _mock_feedback(resume_text: str, job_role: str) -> dict:
    words = resume_text.split()
    word_count = len(words)

    text_lower = resume_text.lower()
    skill_hits = sum(1 for s in SKILLS_DATABASE if s in text_lower)

    base = 50
    word_bonus = min(20, word_count // 100)
    skill_bonus = min(20, skill_hits)
    role_bonus = 10 if job_role.lower() in text_lower else 0

    score = min(100, base + word_bonus + skill_bonus + role_bonus)

    strengths = []
    weaknesses = []
    suggestions = []

    if word_count > 300:
        strengths.append("Resume has substantial content and detail.")
    if skill_hits > 8:
        strengths.append("Strong breadth of technical skills detected.")
    if "project" in text_lower or "projects" in text_lower:
        strengths.append("Projects section appears to be present.")
    if "experience" in text_lower:
        strengths.append("Work experience is mentioned.")
    if not strengths:
        strengths.append("Resume was submitted successfully.")

    if word_count < 150:
        weaknesses.append("Resume appears too short; consider expanding.")
    if skill_hits < 4:
        weaknesses.append("Few recognisable technical skills detected.")
    if "education" not in text_lower:
        weaknesses.append("Education section may be missing.")
    if not weaknesses:
        weaknesses.append("No major structural issues detected.")

    suggestions.append(f"Tailor your resume specifically for the {job_role} role.")
    suggestions.append("Quantify achievements (e.g., 'Increased accuracy by 15%').")
    suggestions.append("Add links to your GitHub, portfolio, or LinkedIn.")
    if skill_hits < 10:
        suggestions.append("Expand your skills section with relevant technologies.")

    return {
        "score": score,
        "strengths": strengths,
        "weaknesses": weaknesses,
        "suggestions": suggestions,
    }


# ==============================================================================
# 8. LEARNING TIPS
# ==============================================================================

TIPS = {
    "python": "Build a small project on GitHub — even a CLI tool or data scraper shows hands-on experience.",
    "java": "Complete a course on Coursera or Udemy, then build a simple backend project.",
    "sql": "Practice on LeetCode SQL or Mode Analytics and document your queries.",
    "html": "Build a personal portfolio page to practice semantic HTML.",
    "css": "Clone a real UI to sharpen your CSS.",
    "javascript": "Build an interactive to-do list or weather app using vanilla JS.",
    "machine learning": "Work through the fast.ai Practical Deep Learning course.",
    "ai": "Explore Hugging Face tutorials and experiment with pre-trained models.",
    "data structures": "Practice DSA problems on LeetCode or GeeksforGeeks.",
    "statistics": "Khan Academy's probability & statistics module is a great free resource.",
    "react": "Build a multi-page React app with routing and reusable components.",
    "node.js": "Create a REST API with Express and connect it to MongoDB.",
    "docker": "Containerize one of your projects and write a clean Dockerfile.",
    "aws": "Deploy a simple app or static site using AWS free tier.",
    "figma": "Redesign a real app screen and create a clickable prototype.",
    "tensorflow": "Train a beginner ML model and explain the workflow in your resume."
}


# ==============================================================================
# 9. MAIN ANALYZER
# ==============================================================================

def analyze_resume(job: str, resume_text: str = "", file_path: str = None, openai_api_key: str = "") -> dict:
    if job not in JOBS:
        raise ValueError("Invalid job role selected.")

    if file_path:
        file_text = read_resume(file_path)
        resume_text = f"{resume_text}\n{file_text}".strip()

    if not resume_text.strip():
        raise ValueError("Resume content is empty.")

    required_skills = JOBS[job]["skills"]
    extracted_skills = extract_skills(resume_text)

    gap_result = skill_gap(extracted_skills, required_skills)
    system_score = calculate_score(gap_result)
    ai_feedback = generate_ai_feedback(resume_text, job, openai_api_key)
    combined_score = final_score(system_score, ai_feedback["score"])

    recommendations = []
    for skill in list(gap_result["missing"])[:5]:
        recommendations.append(TIPS.get(skill, f"Learn {skill} and build a small project to demonstrate it."))

    if not recommendations:
        recommendations.append("Your profile looks strong for this role. Tailor your resume keywords and prepare for interviews.")

    return {
        "job": job,
        "required_skills": sorted(required_skills),
        "extracted_skills": sorted(list(extracted_skills)),
        "matching_skills": sorted(list(gap_result["matching"])),
        "missing_skills": sorted(list(gap_result["missing"])),
        "extra_skills": sorted(list(gap_result["extra"])),
        "system_score": system_score,
        "ai_score": ai_feedback["score"],
        "final_score": combined_score,
        "strengths": ai_feedback["strengths"],
        "weaknesses": ai_feedback["weaknesses"],
        "suggestions": ai_feedback["suggestions"],
        "recommendations": recommendations
    }