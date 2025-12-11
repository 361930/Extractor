import re
import os
import json
import requests
import traceback
from typing import Optional, List, Tuple
from pathlib import Path 

# Import OCR and image libraries
import fitz  # PyMuPDF
import docx2txt
import pytesseract
from PIL import Image
import io
import time 

from utils import log_error

# ----------------- Configuration / Workspace -----------------
HOME = Path.home()
WORKSPACE = HOME / "Desktop" / "ResumeParserWorkspace"
OCR_TEMP_DIR = WORKSPACE / "ocr_temp"
OCR_TEMP_DIR.mkdir(parents=True, exist_ok=True)


# --- Constants ---
OLLAMA_GENERATE_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "llama3:8b"

# --- UPDATED PROMPT FOR EXPERIENCE ---
LLAMA_PROMPT_TEMPLATE = """
You are an expert resume parser. Extract the following details from the resume text below:
1. Candidate Name
2. Email Address
3. Phone Number
4. Total Years of Experience (Numeric value only, e.g., "5", "2.5", "0" if fresher. Estimate based on work history if not explicitly stated).

Respond with ONLY a single, valid JSON object using these exact lowercase keys: "name", "email", "phone", "experience".
Do not include keys that are not requested.
If a field is not found, use an empty string "".

Resume Text:
---
{resume_text}
---

JSON Output:
"""


# --- Text Extraction (OCR Helper) ---

def _run_ocr_on_image(image_path_or_bytes):
    """Helper function to run OCR on a single image (from path or bytes)."""
    
    if not hasattr(_run_ocr_on_image, "tesseract_cmd_path"):
        print("First-time OCR call: Searching for Tesseract-OCR...")
        _run_ocr_on_image.tesseract_cmd_path = None
        
        TESSERACT_PATHS = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"/usr/bin/tesseract", # Linux
            r"/usr/local/bin/tesseract" # Mac
        ]
        
        for path in TESSERACT_PATHS:
            if os.path.exists(path):
                print(f"Found Tesseract-OCR at: {path}")
                _run_ocr_on_image.tesseract_cmd_path = path
                break
        
        if _run_ocr_on_image.tesseract_cmd_path is None:
             print(f"Warning: Tesseract-OCR not found. OCR may fail.")

    if _run_ocr_on_image.tesseract_cmd_path:
        pytesseract.tesseract_cmd = _run_ocr_on_image.tesseract_cmd_path
        
    try:
        img = Image.open(image_path_or_bytes)
        text = pytesseract.image_to_string(img)
        return text
    except Exception as ocr_err:
        log_error(f"OCR failed for image: {ocr_err}")
        return ""


def extract_text(file_path: str) -> Tuple[str, List[str]]:
    """
    Extract text using robust libraries with OCR fallback for PDFs and DOCX.
    """
    text = ""
    lines: List[str] = []
    file_lower = file_path.lower()
    
    try:
        if file_lower.endswith('.pdf'):
            print(f"Extracting with PyMuPDF (fitz) from: {file_path}...")
            text_parts = []
            with fitz.open(file_path) as doc:
                for page_num, page in enumerate(doc):
                    page_text = page.get_text()
                    if page_text:
                        text_parts.append(page_text)
                    
                    image_list = page.get_images(full=True)
                    if image_list:
                        print(f"Page {page_num+1} has {len(image_list)} images. Running OCR...")
                        for img_index, img in enumerate(image_list):
                            try:
                                xref = img[0]
                                base_image = doc.extract_image(xref)
                                image_bytes = base_image["image"]
                                ocr_text = _run_ocr_on_image(io.BytesIO(image_bytes))
                                if ocr_text:
                                    text_parts.append(ocr_text)
                            except Exception:
                                pass
                                
            text = "\n".join(text_parts)
            
        elif file_lower.endswith('.docx'):
            print(f"Extracting with docx2txt from: {file_path}...")
            unique_img_folder = OCR_TEMP_DIR / f"{Path(file_path).stem}_{int(time.time())}"
            unique_img_folder.mkdir(parents=True, exist_ok=True)
            
            text = docx2txt.process(file_path, str(unique_img_folder))
            
            image_text_parts = []
            for img_name in os.listdir(unique_img_folder):
                img_path = unique_img_folder / img_name
                image_text_parts.append(_run_ocr_on_image(str(img_path)))
            
            text = "\n".join(image_text_parts) + "\n" + text
            try:
                for img_name in os.listdir(unique_img_folder):
                    os.remove(unique_img_folder / img_name)
                os.rmdir(unique_img_folder)
            except Exception:
                pass
        else:
            return "", []

        if not text or not text.strip():
            return "", []
            
        lines = text.split('\n')

    except Exception as e:
        log_error(f"Text extraction error for {file_path}: {e}")
        return "", []
        
    cleaned_lines = [re.sub(r'\s+', ' ', line).strip() for line in lines if line.strip()]
    cleaned_text = "\n".join(cleaned_lines)
    
    return cleaned_text[:3500], cleaned_lines


# --- Ollama Parser Function ---

def _call_ollama(text: str) -> Optional[dict]:
    """Internal function to call the Ollama API."""
    
    safe_text = text.replace("{", "{{").replace("}", "}}")
    prompt = LLAMA_PROMPT_TEMPLATE.format(resume_text=safe_text)
    
    payload = {
        "model": MODEL_NAME,
        "prompt": prompt,
        "format": "json", 
        "stream": False
    }

    try:
        print(f"Sending text to {MODEL_NAME}...")
        response = requests.post(OLLAMA_GENERATE_URL, json=payload, timeout=120)
        response.raise_for_status()
        
        response_data = response.json()
        json_string = response_data.get('response')
        
        if not json_string: return None
            
        # Robust JSON extraction
        json_start = json_string.find('{')
        json_end = json_string.rfind('}')
        
        if json_start == -1 or json_end == -1: return None
            
        json_block = json_string[json_start : json_end + 1]
        parsed_json = json.loads(json_block)
        
        # Helper to handle capitalization variants
        def get_val(key):
            return parsed_json.get(key) or parsed_json.get(key.capitalize()) or ""

        return {
            "name": get_val("name"),
            "email": get_val("email"),
            "phone": get_val("phone"),
            "experience": get_val("experience") # New Field
        }
        
    except Exception as e:
        log_error(f"Ollama Error: {e}")
        return None


def parse_resume(file_path: str, nlp_model_unused=None) -> Optional[dict]:
    """
    Parse resume file using the Llama 3 model via Ollama.
    """
    try:
        text, lines = extract_text(file_path)
        if not text or not text.strip():
            return None

        parsed_data = _call_ollama(text)
        
        if parsed_data is None:
            return None

        return {
            "Name": parsed_data.get("name", ""),
            "Email": parsed_data.get("email", ""),
            "Phone": parsed_data.get("phone", ""),
            "Experience": parsed_data.get("experience", "0"), # Default to 0 if missing
            "ResumePath": os.path.abspath(file_path),
            "TextSnippet": text[:500]
        }
        
    except Exception as e:
        log_error(f"parse_resume exception for {file_path}: {e}")
        return None
