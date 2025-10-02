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

# NLP loader
def load_nlp_prefer_transformer(preferred_names=("en_core_web_trf", "en_core_web_sm")):
    """
    Try to load spaCy models by name. This works if they are installed as packages.
    """
    for model_name in preferred_names:
        try:
            nlp = spacy.load(model_name)
            return nlp, model_name
        except OSError:
            # This model is not installed, try the next one
            continue

    # If no models could be loaded, raise an error
    raise OSError(
        f"Could not load any of the preferred spaCy models: {', '.join(preferred_names)}. "
        f"Please install one by running, for example: python -m spacy download en_core_web_sm"
    )