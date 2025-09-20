📄 Resume Parser App (Offline HR Tool)

An offline desktop application to parse resumes (PDF / DOCX), extract useful information (name, email, phone, skills, experience), and save it directly into an Excel database.

Designed for HR teams & recruiters, it requires no internet connection and keeps all resumes + Excel files in a dedicated workspace folder on your Desktop for easy access.

✨ Features

GUI-based (no command line needed, built with Tkinter).

Resume parsing using NLP (spaCy) + regex for:

Name

Email

Phone number

Skills (from a customizable skills.json list)

Work experience (years)

Excel integration:

Create/select active Excel file

Auto-save parsed results

Switch between multiple Excel files

Export visible data to CSV

Duplicate email check (within configurable time window)

Column management:

Add / remove / reorder columns in Excel

Customize headers to fit your workflow

Workspace folder:

Automatically created on Desktop (ResumeParserWorkspace)

Keeps Excel and logs organized in one place

Error logging: all runtime issues logged under logs/errors.log.

📂 Project Structure
ResumeParserApp/
├── app.py                 # Main GUI
├── parser.py              # Resume parsing logic
├── excel_handler.py       # Excel utilities
├── utils.py               # Config + logging + NLP loader
├── skills.json            # Skills keyword list
├── config.json            # Auto-created settings
├── requirements.txt       # Python dependencies
├── README.md              # This file
└── logs/
    └── errors.log         # Error log (auto-created)

🚀 Getting Started
1. Install Dependencies

From the project folder, run:

pip install -r requirements.txt

2. Download spaCy Model

This app prefers the small English model:

python -m spacy download en_core_web_sm


(Optional: for more accuracy, you can also download en_core_web_trf, but it requires PyTorch and will be much larger.)

3. Run the App
python app.py


The GUI will launch.

📖 Usage

Start the app → the main window opens.

Create or select Excel file → all parsed resumes are stored there.

Upload resumes:

Single resume → Upload Resume

Multiple resumes → Upload Multiple

View parsed data → shows in table view.

Edit Excel columns → customize headers.

Export to CSV if needed.

🛠 Notes

Workspace: By default, all files are kept under
~/Desktop/ResumeParserWorkspace/

Duplicates: The app can prevent duplicate candidates based on email within a chosen time window.

Logs: Any error messages are written to logs/errors.log for troubleshooting.

📦 Packaging (Optional)

To make a standalone .exe:

pyinstaller --noconsole --onefile app.py


⚠️ If you include large spaCy models (like en_core_web_trf), the EXE can be several hundred MB.

📌 To-Do / Future Improvements

Improve experience extraction (handle ranges like 2019–2022).

Add advanced filters/search in GUI.

Support scanned PDFs with OCR (e.g., Tesseract).

Add direct resume file organization into folders.
