# Capstone Project: Personal Labs Tracker (Vibe Coding with Google)

A local Python-based intelligent web application that automates the tracking of blood test results over time. The system parses multi-format medical lab data (PDFs, images, text) using Gemini 3.5 Flash, dynamically establishes a grouped tracking template from visual design layouts, and automatically builds or updates a chronological trend-tracking Excel spreadsheet.

---

## 🚀 Key Features & Architectural Highlights

- **Dynamic Template Baseline Generation**: Analyzes 6 separate panel images to map and combine a long list of parameters into a unified master Excel tracking spreadsheet.
- **Multimodal Data Extraction**: Leverages Gemini 3.5 Flash via a single consolidated pipeline to parse incoming text, images, or PDFs into structured data payloads.
- **Chronological Columnar Appending**: Automatically shifts rightward on the spreadsheet to insert new tests under a custom 3-line stacked chronological header (`[Year]\n[MonthDate]\n[Time]`) with forced text-wrapping.
- **Intelligent Parameter Normalization**: Implements an custom LLM mapping rule to unify disparate nomenclature (e.g., standardizing `Hgb` and `Hb` into `Hemoglobin`) to prevent duplicate entry columns.
- **Automated Verification Launcher**: Immediately opens the finalized or updated spreadsheet via the host system's default Excel viewer upon task completion.

---

## 🛠️ Multi-Skill Architecture

Instead of managing fragmented, low-level scripts, the project orchestrates the application flow via two high-level back-end modular skills linked directly to an interactive frontend layout:

### 1. Skill 1: Multimodal Extraction Parser (`data_parser.py`)
Acts as the text and visual analytical layer. Natively processes input documents (PDFs, txt, jpg, png), matches extracted parameters to standard naming rules, captures the overall test completion timestamp, and outputs a strict JSON file format.

### 2. Skill 2: Structural Excel Manager (`excel_manager.py`)
Acts as the file management and spreadsheet compiler layer. Uses `openpyxl` to run conditional branching:
- **Branch A (File Missing):** Constructs the master spreadsheet template from scratch by evaluating the layout boundaries of the 4 included panel images, organizing rows into structural test sections (e.g., *Complete Blood Count*, *Comprehensive Chemistry*, *Healthy Lipid Panel*, *Inflammation Markers*), and mapping reference boundary bounds (Normal vs. Optimal).
- **Branch B (File Present):** Appends newly discovered lab values directly under a clean, vertical 3-line date header column on the right side of the active tracking page.

---

## 🎨 User Interface (Tailwind CSS Dashboard)
The application utilizes an elegant Tailwind CSS frontend layout containing:
- An input text field for targeted file path destinations (for launching new files or appending old ones).
- A drag-and-drop dotted boundary zone to seamlessly capture uploaded lab file payloads from the desktop.
- Interactive backend status monitoring bars.

---


## 💻 Local Setup & Execution Instructions

1. Extract the contents of the project ZIP folder onto your local drive.
2. Rename `.env.example` to `.env`
3. Inside `.env` file, update value for GEMINI_API_KEYwith your personal API credentials string.
4. Double-click **`run.bat`** to start up the local server pipeline, launch the web dashboard. It will create excel spread sheet if it doesn't exists, or will update the existing one, and open it when it is done.
