import re
import pdfplumber
import docx
import os
import json
from typing import Optional, List, Set
from datetime import datetime

# Regexes (more robust)
# Finds most common email formats
email_regex = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
# Finds common phone formats, including international prefixes and extensions
phone_regex = re.compile(r"[\+\(]?[0-9][\s\-\(\)\.0-9]{7,}[0-9]")

def find_emails(text: str) -> str:
    """Finds all unique email addresses in the text."""
    if not text:
        return ""
    matches = email_regex.findall(text)
    # Return unique, sorted matches joined by a comma
    return ", ".join(sorted(list(set(m.strip() for m in matches))))

def find_phones(text: str) -> str:
    """Finds all unique, cleaned phone numbers in the text."""
    if not text:
        return ""
    matches = phone_regex.findall(text)
    # Clean matches to only include digits and '+'
    cleaned_matches: Set[str] = set()
    for m in matches:
        # Basic cleaning: remove spaces, dots, parens, hyphens
        cleaned = re.sub(r"[\s\.\(\)\-]", "", m)
        if len(cleaned) > 7: # Simple filter for valid-length numbers
            cleaned_matches.add(cleaned)
    
    # Return unique, sorted matches joined by a comma
    return ", ".join(sorted(list(cleaned_matches)))


def extract_text(file_path: str) -> str:
    """
    Extract text from PDF (via pdfplumber) or DOCX.
    Returns empty string if extraction fails.
    """
    text = ""
    file_lower = file_path.lower()
    
    try:
        if file_lower.endswith('.pdf'):
            with pdfplumber.open(file_path) as pdf:
                for p in pdf.pages:
                    page_text = p.extract_text()
                    if page_text:
                        text += page_text + "\n"
        elif file_lower.endswith('.docx'):
            doc = docx.Document(file_path)
            for p in doc.paragraphs:
                text += p.text + "\n"
    except Exception as e:
        # Log this error if possible, but return empty string
        print(f"Error extracting text from {file_path}: {e}") # Simple console log
        return ""
        
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def fallback_name_from_text(text: str) -> str:
    """Heuristic name fallback: first short line with capitalized words."""
    # Look at the first few lines of text
    for line in text.splitlines()[:10]: # Check first 10 lines
        s = line.strip()
        if not s:
            continue
        parts = s.split()
        # heuristic: short line (2-4 words) with each starting capitalized
        # and no obvious email/phone triggers
        if (1 < len(parts) <= 4 and 
            all(p and p[0].isupper() for p in parts) and
            "@" not in s and "http" not in s):
            # Check if it looks like a name (avoids titles like "CURRICULUM VITAE")
            if all(re.match(r"^[A-Z][a-z\-']+$", p, re.IGNORECASE) for p in parts):
                 return s
    return ""

# Main parser
def parse_resume(file_path: str, nlp) -> Optional[dict]:
    """
    Parse resume file and return dict or None on failure.
    Focuses on Name, Email, and Phone.
    """
    try:
        text = extract_text(file_path)
        if not text or not text.strip():
            return None # Skip if no text

        # 1. NER for name if spaCy nlp available
        name = ""
        if nlp:
            try:
                # Process only the first 1000 chars for speed/accuracy of name
                doc = nlp(text[:1000]) 
                for ent in doc.ents:
                    if ent.label_ == "PERSON":
                        # Perform basic validation
                        if len(ent.text.strip().split()) > 1 and len(ent.text.strip()) < 35:
                            name = ent.text.strip()
                            break # Take first valid PERSON
            except Exception as e:
                print(f"SpaCy error: {e}")
                name = "" # Fallback

        # 2. Fallback name extraction if spaCy fails or finds nothing
        if not name:
            name = fallback_name_from_text(text)

        # 3. Extract Emails and Phones
        email = find_emails(text)
        phone = find_phones(text)

        # If all fields are empty, it's likely a bad parse
        if not name and not email and not phone:
            return None

        return {
            "Name": name or "",
            "Email": email or "",
            "Phone": phone or "",
            "ResumePath": os.path.abspath(file_path),
            "TextSnippet": text[:800] # Keep snippet for debug
        }
    except Exception as e:
        try:
            from utils import log_error
            log_error(f"parse_resume exception for {file_path}: {e}")
        except Exception:
            print(f"parse_resume exception for {file_path}: {e}")
        return None

