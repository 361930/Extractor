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

# --- Tesseract-OCR Configuration has been MOVED ---
# We no longer set the path here, as it doesn't work across threads.
# The logic is now inside the `_run_ocr_on_image` function.


# ----------------- Configuration / Workspace -----------------
HOME = Path.home()
WORKSPACE = HOME / "Desktop" / "ResumeParserWorkspace"
OCR_TEMP_DIR = WORKSPACE / "ocr_temp"
OCR_TEMP_DIR.mkdir(parents=True, exist_ok=True)


# --- Constants ---
OLLAMA_GENERATE_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "llama3:8b"

LLAMA_PROMPT_TEMPLATE = """
You are an expert resume parser. Your *only* job is to extract JSON data.
You will be given resume text. Extract the candidate's name, email, and phone number.

Respond with ONLY a single, valid JSON object using these exact lowercase keys: "name", "email", "phone".
DO NOT add any conversational text, explanations, or apologies.
Your entire response MUST be the JSON object.

If a field is not found, use an empty string "".

Resume Text:
---
{resume_text}
---

JSON Output:
"""


# --- Text Extraction (NOW WITH ROBUST OCR) ---

def _run_ocr_on_image(image_path_or_bytes):
    """Helper function to run OCR on a single image (from path or bytes)."""
    
    # --- NEW FIX: Set the path *inside* the function ---
    # This ensures it's set in the correct thread.
    
    # We use a static function attribute to find the path only ONCE.
    if not hasattr(_run_ocr_on_image, "tesseract_cmd_path"):
        print("First-time OCR call: Searching for Tesseract-OCR...")
        _run_ocr_on_image.tesseract_cmd_path = None # Set a static variable
        
        TESSERACT_PATHS = [
            # --- ADD YOUR CUSTOM PATH HERE IF NEEDED ---
            # r"C:\Your\Custom\Path\Tesseract-OCR\tesseract.exe",
            # ---------------------------------
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        ]
        
        for path in TESSERACT_PATHS:
            if os.path.exists(path):
                print(f"Found Tesseract-OCR at: {path}")
                _run_ocr_on_image.tesseract_cmd_path = path # Store it
                break
        
        if _run_ocr_on_image.tesseract_cmd_path is None:
             print(f"Warning: Tesseract-OCR not found in default locations: {TESSERACT_PATHS}")
             print("Hoping it's in your system PATH...")
             # We let it be None and let pytesseract try to find it.

    # Now, set the path for the *current* thread
    # This is the line that fixes the bug
    if _run_ocr_on_image.tesseract_cmd_path:
        pytesseract.tesseract_cmd = _run_ocr_on_image.tesseract_cmd_path
    # --- END NEW FIX ---
        
    try:
        img = Image.open(image_path_or_bytes)
        # Run OCR
        text = pytesseract.image_to_string(img)
        return text
    except pytesseract.TesseractNotFoundError:
        log_error("TESSERACT_NOT_FOUND: 'tesseract.exe' was not found.")
        print("CRITICAL ERROR: 'tesseract.exe' not found. Please install Tesseract-OCR.")
        return "" 
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
                    
                    # --- FIX 1: ALWAYS get the normal text ---
                    page_text = page.get_text()
                    if page_text:
                        text_parts.append(page_text)
                    
                    # --- FIX 2: ALWAYS get the images and run OCR ---
                    image_list = page.get_images(full=True)
                    if image_list:
                        print(f"Page {page_num+1} has {len(image_list)} images. Running OCR on them...")
                        for img_index, img in enumerate(image_list):
                            try:
                                xref = img[0]
                                base_image = doc.extract_image(xref)
                                image_bytes = base_image["image"]
                                
                                # Run OCR on the image bytes
                                ocr_text = _run_ocr_on_image(io.BytesIO(image_bytes))
                                if ocr_text:
                                    print(f"  > OCR found text in image {img_index} on page {page_num+1}")
                                    text_parts.append(ocr_text)
                            except Exception as img_ex:
                                log_error(f"Failed to extract/OCR image {img_index} on page {page_num+1}: {img_ex}")
                                
            text = "\n".join(text_parts)
            
        elif file_lower.endswith('.docx'):
            print(f"Extracting with docx2txt from: {file_path}...")
            
            # This logic is already robust (gets text + images)
            unique_img_folder = OCR_TEMP_DIR / f"{Path(file_path).stem}_{int(time.time())}"
            unique_img_folder.mkdir(parents=True, exist_ok=True)
            
            text = docx2txt.process(file_path, str(unique_img_folder))
            
            image_text_parts = []
            for img_name in os.listdir(unique_img_folder):
                img_path = unique_img_folder / img_name
                print(f"Found image in docx, running OCR on: {img_name}")
                image_text_parts.append(_run_ocr_on_image(str(img_path)))
            
            text = "\n".join(image_text_parts) + "\n" + text
            try:
                for img_name in os.listdir(unique_img_folder):
                    os.remove(unique_img_folder / img_name)
                os.rmdir(unique_img_folder)
            except Exception as e:
                log_error(f"Failed to clean up temp image folder {unique_img_folder}: {e}")
            
        else:
            print(f"Skipping {file_path}: Unsupported file type.")
            return "", []

        if not text or not text.strip():
            print(f"Skipping {file_path}: No text found after extraction and OCR.")
            return "", []
            
        lines = text.split('\n')

    except Exception as e:
        print(f"Error extracting text from {file_path}: {e}")
        log_error(f"Text extraction error for {file_path}: {e}\n{traceback.format_exc()}")
        return "", []
        
    cleaned_lines = [re.sub(r'\s+', ' ', line).strip() for line in lines if line.strip()]
    cleaned_text = "\n".join(cleaned_lines)
    
    return cleaned_text[:3000], cleaned_lines


# --- Ollama Parser Function (Fixes JSON and KeyErrors) ---

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
        print(f"Sending resume text to {MODEL_NAME} via Ollama...")
        response = requests.post(OLLAMA_GENERATE_URL, json=payload, timeout=120)
        response.raise_for_status()
        
        response_data = response.json()
        
        json_string = response_data.get('response')
        if not json_string:
            log_error("Ollama response was empty.")
            return None
            
        json_start = json_string.find('{')
        json_end = json_string.rfind('}')
        
        if json_start == -1 or json_end == -1:
            log_error(f"Ollama response did not contain JSON: {json_string}")
            return None
            
        json_block = json_string[json_start : json_end + 1]
            
        parsed_json = json.loads(json_block)
        
        def get_key(data, key):
            if key in data: return data[key]
            if key.capitalize() in data: return data[key.capitalize()]
            return None

        name_val = get_key(parsed_json, "name")
        email_val = get_key(parsed_json, "email")
        phone_val = get_key(parsed_json, "phone")

        if name_val is None or email_val is None or phone_val is None:
            log_error(f"Ollama's JSON response was missing keys: {json_block}")
            return None
        
        return {
            "name": name_val,
            "email": email_val,
            "phone": phone_val
        }
        
    except json.JSONDecodeError as e:
        log_error(f"Failed to parse JSON response from Ollama: {e}")
        try:
            log_error(f"Raw response was: {json_string}")
        except Exception:
            pass
        return None
    except requests.exceptions.RequestException as e:
        log_error(f"Ollama API request failed (e.g., timeout): {e}")
        return None
    except Exception as e:
        log_error(f"Error in _call_ollama: {e}\n{traceback.format_exc()}")
        return None


# --- Main Parser Function (Fixes NameError: 'parsed_') ---
def parse_resume(file_path: str, nlp_model_unused) -> Optional[dict]:
    """
    Parse resume file using the Llama 3 model via Ollama.
    """
    try:
        text, lines = extract_text(file_path)
        if not text or not text.strip():
            return None

        parsed_data = _call_ollama(text)
        
        if parsed_data is None:
            log_error(f"Ollama parsing failed for {file_path}.")
            return None

        # Format the data for the Excel sheet
        return {
            "Name": parsed_data.get("name", ""),
            "Email": parsed_data.get("email", ""),
            "Phone": parsed_data.get("phone", ""),
            "ResumePath": os.path.abspath(file_path),
            "TextSnippet": text[:800] # Keep snippet for debug
        }
        
    except Exception as e:
        log_error(f"parse_resume exception for {file_path}: {e}\n{traceback.format_exc()}")
        return None
