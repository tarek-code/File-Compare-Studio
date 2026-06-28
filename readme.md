# File Compare Studio 🚀

Welcome to **File Compare Studio**, a suite of modern, lightweight, and high-performance Python utilities designed to compare, filter, and analyze CSV and Excel (`.xlsx`, `.xls`) files. The tools read files as text to avoid floating-point discrepancies, apply dynamic filters, and identify unique and common rows without altering your original source files.

Developed & Copyrighted by **Tarek (Unlimited)**. All Rights Reserved © 2026.

---

## 📬 Contact & Communication
If you have any questions, feedback, bug reports, feature requests, or want to collaborate, feel free to reach out directly:
*   **Email:** [tarekadel314@gmail.com](mailto:tarekadel314@gmail.com)
*   **Developer:** Tarek (Unlimited)

---

## 🛠️ Environment Setup & Installation

To run File Compare Studio, follow these steps to set up a clean Python virtual environment and install the required dependencies.

### Step 1: Create a Virtual Environment (`venv`)
Creating a virtual environment ensures that the project dependencies do not conflict with other Python packages installed globally on your machine.

**On Windows:**
```powershell
python -m venv .venv
```

### Step 2: Activate the Virtual Environment
Activate the environment so that any python commands or package installations occur within the virtual sandbox.

**On Windows (Command Prompt / PowerShell):**
```powershell
.venv\Scripts\activate
```

### Step 3: Install Required Dependencies
With the virtual environment active, run the following command to install the required libraries listed in `requirements.txt`:
```powershell
pip install -r requirements.txt
```

*Note: The requirements include:*
*   `pandas` (For high-speed table operations and file parsing)
*   `openpyxl` (For reading/writing Excel `.xlsx` spreadsheets)
*   `customtkinter` (For the modern Dark/Light theme desktop user interface)
*   `packaging` (For version checking)

---

## 📂 Project Structure & Folders

The project is divided into three main versions, each targeting a specific workflow or environment. Below is a breakdown of the directories and when to use them:

```text
File Compare Studio/
│
├── CMD Version/                # Command-Line Interface (CLI) version
│   └── File-Compare-Studio-CMD-Light.py
│
├── Lighte GUI/                 # Modern, lightweight dual-file comparison GUI
│   └── File-Compare-Studio-Light.py
│
├── Full Version/               # Advanced workspace GUI supporting multiple files
│   ├── File-Compare-Studio.py
│   ├── NA_129.csv
│   ├── NA_190.csv
│   └── noms.xlsx
│
├── requirements.txt            # Package dependencies
└── readme.md                   # This documentation file
```

---

### 1. 💻 CMD Version (`CMD Version/`)
*   **Core File:** `CMD Version/File-Compare-Studio-CMD-Light.py`
*   **When to Use:**
    *   You are working in a headless environment (like a server, SSH session, or Linux terminal without a display server).
    *   You prefer using terminal inputs and want a fast, interactive keyboard-driven workflow.
    *   You want to run quick, automated, or script-assisted comparisons.
*   **How to Run:**
    ```powershell
    python "CMD Version/File-Compare-Studio-CMD-Light.py"
    ```

---

### 2. ⚡ Light GUI Version (`Lighte GUI/`)
*   **Core File:** `Lighte GUI/File-Compare-Studio-Light.py`
*   **When to Use:**
    *   You want a modern, visually stunning desktop window but only need to compare **exactly two files** at a time.
    *   You want to filter records dynamically before comparing and export separate files for:
        1. Rows unique to File 1.
        2. Rows unique to File 2.
        3. Rows common to both files.
    *   It is optimized for quick, daily comparisons with a clean and minimal layout.
*   **How to Run:**
    ```powershell
    python "Lighte GUI/File-Compare-Studio-Light.py"
    ```

---

### 3. 🌟 Full Version (`Full Version/`)
*   **Core File:** `Full Version/File-Compare-Studio.py`
*   **When to Use:**
    *   You need a robust, **workspace-based application** where you can load **multiple files** (more than 2) simultaneously.
    *   You require advanced comparisons, asynchronous processing (with background progress bars to prevent window freezing on large files), and custom columns mapping.
    *   You want to manage a full workspace session, configure complex keep-filters, and choose exactly which specific export types you want to generate.
*   **How to Run:**
    ```powershell
    python "Full Version/File-Compare-Studio.py"
    ```

---

## 🔒 Data Integrity & Processing Safety
*   **Read-Only Operations:** All files loaded into File Compare Studio are read as strictly read-only. The application never modifies, overwrites, or alters your source spreadsheets.
*   **Strict String Processing:** To ensure accuracy, the tool parses all columns as raw text strings. This avoids common Excel issues where numerical values (like ID numbers with leading zeros, e.g., `00123`) get truncated or modified during processing.
