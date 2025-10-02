import re
import pdfplumber
import docx
import os
import json
from typing import Optional
from datetime import datetime

# Skills list loader (fallback small list)
SKILLS_FILE = "skills.json"
if os.path.exists(SKILLS_FILE):
    try:
        with open(SKILLS_FILE, "r", encoding="utf-8") as f:
            SKILLS_LIST = [s.lower() for s in json.load(f)]
    except Exception:
        SKILLS_LIST = ["python", "excel", "sql", "java", "aws", "linux", "docker"]
else:
    SKILLS_LIST = ["python", "excel", "sql", "java", "aws", "linux", "docker"]

# Regexes
email_regex = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
phone_regex = re.compile(r"\+?[0-9][0-9\-\s()]{6,}[0-9]")

# matches "3 years", "3.5 years", "4 yrs", "2+ years", "5-year"
simple_exp_re = re.compile(r"(\d+(?:\.\d+)?)(?:\+)?\s*(?:-year|years|year|yrs|yrs\.)\b", re.I)
# matches "Total experience: 5 years" etc
total_exp_re = re.compile(r"(?:total\s+experience[:\s]*|experience[:\s]*)(\d+(?:\.\d+)?)\s*(?:years|yrs|year)\b", re.I)
# matches date ranges like "Jan 2019 - Mar 2022" or "2018 - 2021" or "2018–2021"
date_range_re = re.compile(
    r"(?P<start>(?:\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b\.?\s*)?(?:\d{4}))\s*(?:[-–—to]{1,3})\s*(?P<end>(?:present|current|now)|(?:\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b\.?\s*)?(?:\d{4}))",
    re.I
)

MONTHS = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"sept":9,"oct":10,"nov":11,"dec":12}

def find_email(text: str) -> str:
    if not text:
        return ""
    m = email_regex.search(text)
    return m.group(0).strip() if m else ""

def find_phone(text: str) -> str:
    if not text:
        return ""
    m = phone_regex.search(text)
    return m.group(0).strip() if m else ""

def _parse_year_from_token(token: str) -> int | None:
    if not token:
        return None
    token = token.strip().lower().replace('.', '')
    parts = token.split()
    if len(parts) == 1:
        if parts[0].isdigit() and len(parts[0]) == 4:
            return int(parts[0])
        return None
    try:
        month = parts[0][:3]
        if month in MONTHS and parts[-1].isdigit() and len(parts[-1]) == 4:
            return int(parts[-1])
    except Exception:
        return None
    if parts[-1].isdigit() and len(parts[-1]) == 4:
        return int(parts[-1])
    return None

def find_experience(text: str) -> str:
    if not text:
        return ""
    m = total_exp_re.search(text)
    if m:
        try:
            val = float(m.group(1))
            return f"{int(val) if val.is_integer() else val} years"
        except Exception:
            pass
    simple_matches = [float(m.group(1)) for m in simple_exp_re.finditer(text)]
    if simple_matches:
        val = max(simple_matches)
        return f"{int(val) if val.is_integer() else val} years"

    years_candidates = []
    for dm in date_range_re.finditer(text):
        start_year = _parse_year_from_token(dm.group("start"))
        end_tok = dm.group("end")
        end_year = datetime.now().year if isinstance(end_tok, str) and end_tok.strip().lower() in ("present","current","now") else _parse_year_from_token(end_tok)
        if start_year and end_year and end_year >= start_year:
            years_candidates.append(end_year - start_year)

    if years_candidates:
        return f"{max(years_candidates)} years"
    return ""

def find_skills(text: str) -> str:
    if not text:
        return ""
    lower = text.lower()
    return ", ".join(sorted(list(set(skill for skill in SKILLS_LIST if skill in lower))))

def extract_text(file_path: str) -> str:
    text = ""
    try:
        if file_path.lower().endswith('.pdf'):
            with pdfplumber.open(file_path) as pdf:
                for p in pdf.pages:
                    page_text = p.extract_text()
                    if page_text:
                        text += page_text + "\n"
        elif file_path.lower().endswith('.docx'):
            doc = docx.Document(file_path)
            for p in doc.paragraphs:
                text += p.text + "\n"
    except Exception as e:
        log_error(f"Text extraction failed for {file_path}: {e}")
        return ""
    return text

def fallback_name_from_text(text: str) -> str:
    for line in text.splitlines():
        s = line.strip()
        if not s: continue
        parts = s.split()
        if 1 < len(parts) <= 4 and all(p and p[0].isupper() for p in parts):
            return s
    return ""

def parse_resume(file_path: str, nlp) -> Optional[dict]:
    try:
        text = extract_text(file_path)
        if not text or not text.strip():
            return None

        name = ""
        try:
            doc = nlp(text)
            for ent in doc.ents:
                if ent.label_ == "PERSON":
                    name = ent.text.strip()
                    break
        except Exception:
            name = ""

        if not name:
            name = fallback_name_from_text(text)

        return {
            "Name": name,
            "Email": find_email(text),
            "Phone": find_phone(text),
            "Skills": find_skills(text),
            "Experience": find_experience(text),
            "ResumePath": os.path.abspath(file_path),
            "TextSnippet": text[:800]
        }
    except Exception as e:
        log_error(f"parse_resume exception for {file_path}: {e}")
        return None