📄 Resume Parser App (Contact Details Extractor)
An offline desktop application designed to bulk-parse resumes (PDF / DOCX), extract key contact information (Name, Email, Phone), and organize it into a structured Excel database.
This tool is streamlined for speed and accuracy, making it ideal for recruiters who need to quickly build a contact list from a large repository of candidate files.
✨ Key Features
 * Streamlined Extraction: Focuses strictly on essential contact details:
   * Name (using NLP + Heuristics)
   * Email (extracts all unique emails found in the document)
   * Phone (extracts all unique phone numbers)
 * Robust Parsing: Handles various resume formats and layouts. Finds multiple contact numbers or emails if present.
 * Bulk Processing:
   * Upload Files: Select multiple specific files at once.
   * Upload Folder: New! Select a folder to recursively scan and process all PDF/DOCX files inside it.
 * Excel Integration:
   * Auto-Organized: Data is saved with a Serial Number (S.No.) for easy counting.
   * Duplicate Protection: Built-in check to prevent adding the same candidate (by email) within a configurable time window (default: 30 days).
   * Original File Tracking: Captures the original filename alongside the parsed data.
 * Modern GUI: Clean, user-friendly interface built with Tkinter (styled with modern fonts and colors).
 * Offline & Secure: Runs entirely on your local machine. No data is uploaded to the cloud.
📂 Output Structure
The application creates a ResumeParserWorkspace on your Desktop containing:
 * excel/: Stores the database (e.g., resumes_data.xlsx).
 * resumes/: Stores a copy of the processed resume files.
 * logs/: Error logs for troubleshooting.
Excel Columns:
 * S.No. (Serial Number)
 * Name
 * Email
 * Phone
 * OriginalFile
 * DateApplied
 * ResumePath
🚀 Getting Started
1. Install Dependencies
Ensure you have Python installed. Then, install the required libraries:
pip install -r requirements.txt

2. Download spaCy Model
The app uses a lightweight English model for name recognition:
python -m spacy download en_core_web_sm

(Optional: For higher accuracy, you can install en_core_web_trf, though it is larger.)
3. Run the App
python app.py

📖 Usage Guide
 * Launch: Run the script to open the GUI.
 * Excel Setup: The app automatically creates a default Excel file. You can use the "Create / Select Excel" button to switch files.
 * Upload:
   * Click Upload Files to select specific documents.
   * Click Upload Folder to process an entire directory of resumes.
 * Review: As files are processed, they appear in the table view.
 * Export: Click Export Visible → CSV to save the current table view as a CSV file.
 * Open Excel: Click Open Excel to view the full database in Microsoft Excel.
🛠 Technical Details
 * Language: Python 3
 * GUI: Tkinter
 * Parsing: pdfplumber (PDF text), python-docx (Word text), spaCy (Name NER), Regex (Email/Phone).
 * Excel Engine: openpyxl
📦 Packaging (exe)
To create a standalone executable file for distribution:
pyinstaller app.spec

Note: The generated .exe will be in the dist folder.
📄 License
MIT License - See LICENSE file for details.

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

