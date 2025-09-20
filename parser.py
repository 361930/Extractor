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
experience_regex = re.compile(r"(\d{1,2})\s*(?:\+)?\s*(?:years|year|yrs|yr)\b", re.I)

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

MONTHS = {
    "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
    "jul":7,"aug":8,"sep":9,"sept":9,"oct":10,"nov":11,"dec":12
}

# matches "3 years", "3.5 years", "4 yrs", "2+ years", "5-year"
simple_exp_re = re.compile(r"(\d+(?:\.\d+)?)(?:\+)?\s*(?:-year|years|year|yrs|yrs\.)\b", re.I)
# matches "Total experience: 5 years" etc
total_exp_re = re.compile(r"(?:total\s+experience[:\s]*|experience[:\s]*)(\d+(?:\.\d+)?)\s*(?:years|yrs|year)\b", re.I)
# matches date ranges like "Jan 2019 - Mar 2022" or "2018 - 2021" or "2018–2021"
date_range_re = re.compile(
    r"(?P<start>(?:\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b\.?\s*)?(?:\d{4}))\s*(?:[-–—to]{1,3})\s*(?P<end>(?:present|current|now)|(?:\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b\.?\s*)?(?:\d{4}))",
    re.I
)

def _parse_year_from_token(token: str) -> int | None:
    if not token:
        return None
    token = token.strip().lower().replace('.', '')
    # token may be "Jan 2019" or "2019"
    parts = token.split()
    if len(parts) == 1:
        # maybe "2019"
        if parts[0].isdigit() and len(parts[0]) == 4:
            return int(parts[0])
        return None
    # if month present
    try:
        month = parts[0][:3]
        if month in MONTHS and parts[-1].isdigit() and len(parts[-1]) == 4:
            return int(parts[-1])
    except Exception:
        return None
    # fallback: if last token is year
    if parts[-1].isdigit() and len(parts[-1]) == 4:
        return int(parts[-1])
    return None

def find_experience(text: str) -> str:
    """
    Return human-friendly experience string if found, otherwise empty string.
    Priority:
    1) explicit "Total experience" like lines
    2) simple "3 years" pattern
    3) date ranges -> compute approximate years
    """
    if not text:
        return ""

    # 1) total experience explicit
    m = total_exp_re.search(text)
    if m:
        try:
            val = float(m.group(1))
            # return "X years" formatted
            if val.is_integer():
                return f"{int(val)} years"
            return f"{val} years"
        except Exception:
            pass

    # 2) simple "X years" anywhere - prefer largest match (someone might list multiple)
    simple_matches = [float(m.group(1)) for m in simple_exp_re.finditer(text)]
    if simple_matches:
        # choose the maximum (e.g., "5 years experience" vs "1 year internship")
        val = max(simple_matches)
        if float(val).is_integer():
            return f"{int(val)} years"
        return f"{val} years"

    # 3) date ranges: try to compute
    dr_matches = list(date_range_re.finditer(text))
    years_candidates = []
    for dm in dr_matches:
        start_tok = dm.group("start")
        end_tok = dm.group("end")
        start_year = _parse_year_from_token(start_tok)
        if isinstance(end_tok, str) and end_tok.strip().lower() in ("present","current","now"):
            end_year = datetime.now().year
        else:
            end_year = _parse_year_from_token(end_tok)
        if start_year and end_year:
            delta = end_year - start_year
            if delta >= 0:
                years_candidates.append(delta)

    if years_candidates:
        val = max(years_candidates)
        return f"{val} years"

    return ""

def find_skills(text: str) -> str:
    if not text:
        return ""
    lower = text.lower()
    found = []
    for skill in SKILLS_LIST:
        if skill in lower and skill not in found:
            found.append(skill)
    return ", ".join(found)

def extract_text(file_path: str) -> str:
    """
    Extract text from PDF (via pdfplumber) or DOCX.
    Returns empty string if extraction fails.
    """
    text = ""
    if file_path.lower().endswith('.pdf'):
        try:
            with pdfplumber.open(file_path) as pdf:
                for p in pdf.pages:
                    page_text = p.extract_text()
                    if page_text:
                        text += page_text + "\n"
        except Exception:
            return ""
    elif file_path.lower().endswith('.docx'):
        try:
            doc = docx.Document(file_path)
            for p in doc.paragraphs:
                text += p.text + "\n"
        except Exception:
            return ""
    return text




def find_email(text: str) -> Optional[str]:
    m = email_regex.search(text)
    return m.group(0).strip() if m else ""




def find_phone(text: str) -> Optional[str]:
    m = phone_regex.search(text)
    return m.group(0).strip() if m else ""




def find_experience(text: str) -> Optional[str]:
    m = experience_regex.search(text)
    return (m.group(0).strip()) if m else ""




def find_skills(text: str) -> str:
    lower = text.lower()
    found = []
    for skill in SKILLS_LIST:
        if skill in lower and skill not in found:
         found.append(skill)
    return ", ".join(found)




# Main parser


def parse_resume(file_path: str, nlp) -> Optional[dict]:
    """
    Parse resume file and return dict or None on failure.
    Uses helper functions above for email/phone/skills.
    """
    try:
        # If you have an extract_text function, use it; otherwise read docx/pdf here.
        text = extract_text(file_path)
        if not text or not text.strip():
            return None

        # NER for name if spaCy nlp available
        name = ""
        try:
            doc = nlp(text) if nlp else None
            if doc:
                for ent in doc.ents:
                    if ent.label_ == "PERSON":
                        name = ent.text.strip()
                        break
        except Exception:
            # if spaCy fails for any reason, fallback to heuristics below
            name = ""

        # fallback: first short line with capitalized words
        if not name:
            for line in text.splitlines():
                s = line.strip()
                if not s:
                    continue
                parts = s.split()
                if 1 < len(parts) <= 4 and all(p[0].isupper() for p in parts if p):
                    name = s
                    break

        email = find_email(text)
        phone = find_phone(text)
        experience = find_experience(text)
        skills = find_skills(text)
        experience = find_experience(text)


        return {
            "Name": name or "",
            "Email": email or "",
            "Phone": phone or "",
            "Skills": skills or "",
            "Experience": experience or "",
            "ResumePath": os.path.abspath(file_path),
            "TextSnippet": text[:800]
        }
    except Exception as e:
        # don't raise — return None so caller (app) can show a friendly message
        # log if you have log_error function (optional)
        try:
            from utils import log_error
            log_error(f"parse_resume exception for {file_path}: {e}")
        except Exception:
            pass
        return None

def fallback_name_from_text(text: str) -> str:
    """Heuristic name fallback: first short line with capitalized words."""
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        parts = s.split()
        # heuristic: short line (2-4 words) with each starting capitalized
        if 1 < len(parts) <= 4 and all(p and p[0].isupper() for p in parts):
            return s
    return ""
