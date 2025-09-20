import os
import json
import traceback
import sys
import spacy




CONFIG_FILE = "config.json"
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "errors.log")

DEFAULT_CONFIG = {
    "active_excel": "resumes_data.xlsx",
    "duplicate_check_enabled": True,
    "duplicate_days": 30,
    "last_used_model": "en_core_web_sm"
}

def ensure_dirs():
    os.makedirs(LOG_DIR, exist_ok=True)

def log_error(msg: str):
    ensure_dirs()
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n" + ("-"*80) + "\n")





# Config helpers
def load_config():
    if not os.path.exists(CONFIG_FILE):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log_error("Failed to load config: " + str(e) + "\n" + traceback.format_exc())
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

def save_config(config: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        log_error("Failed to save config: " + str(e) + "\n" + traceback.format_exc())



# NLP loader (tries transformer model first, falls back)
import spacy



def load_nlp_prefer_transformer(preferred_names=("en_core_web_trf", "en_core_web_sm")):
    """
    Try to load spaCy model folder from preferred_names in app dir (or _MEIPASS when frozen).
    Returns (nlp, model_name) or raises OSError if none found.
    """
    base_search_paths = [os.getcwd()]
    if getattr(sys, "frozen", False):
        base_search_paths.insert(0, getattr(sys, "_MEIPASS", os.getcwd()))

    for model_name in preferred_names:
        for base in base_search_paths:
            candidate = os.path.join(base, model_name)
            if os.path.isdir(candidate):
                # attempt to load by path
                try:
                    nlp = spacy.load(candidate)
                    return nlp, model_name
                except Exception:
                    # try by model name (package)
                    try:
                        nlp = spacy.load(model_name)
                        return nlp, model_name
                    except Exception:
                        continue
    # if we reach here, none could be loaded
    raise OSError("No spaCy model found in app directory. Place model folder 'en_core_web_sm' or 'en_core_web_trf' next to the EXE.")
