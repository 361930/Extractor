import os
import json
import traceback
import sys
import requests # New import

CONFIG_FILE = "config.json"
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "errors.log")

DEFAULT_CONFIG = {
    "active_excel": "resumes_data.xlsx",
    "duplicate_check_enabled": True
}

def ensure_dirs():
    os.makedirs(LOG_DIR, exist_ok=True)

def log_error(msg: str):
    """Logs an error message to the log file."""
    ensure_dirs()
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            # FIX: The old line was too complex.
            # The 'msg' variable already contains the full traceback.
            f.write(f"{msg}\n{'-'*80}\n")
    except Exception as e:
        # If logging fails, print to console as a fallback
        print("--- FATAL LOGGING ERROR ---")
        print(f"Failed to write to log file: {e}")
        print(f"Original message: {msg}")
        print("--- END ---")

# Config helpers
def load_config():
    """Loads config.json, creating it with defaults if it doesn't exist."""
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
    """Saves the config dictionary to config.json."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        log_error("Failed to save config: " + str(e) + "\n" + traceback.format_exc())


# --- NEW: Ollama Connection Checker ---
def check_ollama_connection():
    """
    Checks if the Ollama server is running and has the llama3:8b model.
    Returns (True, "model_name") or (False, "error_message").
    """
    OLLAMA_API_URL = "http://localhost:11434/api/tags"
    MODEL_TO_CHECK = "llama3:8b"
    
    try:
        print("Checking connection to Ollama server at http://localhost:11434...")
        
        # 1. Check if server is running
        response = requests.get(OLLAMA_API_URL, timeout=5)
        response.raise_for_status() # Raise error if status is 4xx or 5xx
        
        print("Ollama server is running.")
        
        # 2. Check if the model is downloaded
        models_data = response.json()
        available_models = [model['name'] for model in models_data.get('models', [])]
        
        if MODEL_TO_CHECK in available_models:
            print(f"Found model: {MODEL_TO_CHECK}")
            return True, MODEL_TO_CHECK
        else:
            error_msg = f"Ollama is running, but model '{MODEL_TO_CHECK}' is missing."
            print(error_msg)
            print("Please run this command in your terminal:")
            print(f"ollama pull {MODEL_TO_CHECK}")
            return False, error_msg

    except requests.exceptions.ConnectionError:
        error_msg = "Could not connect to Ollama."
        print(error_msg)
        print("Please make sure the Ollama application is running.")
        return False, error_msg
    except Exception as e:
        error_msg = f"An error occurred: {e}"
        print(error_msg)
        log_error(f"Ollama check failed: {e}\n{traceback.format_exc()}")
        return False, error_msg
