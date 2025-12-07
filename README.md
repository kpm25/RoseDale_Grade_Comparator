# RoseDale_Grade_Comparator
Python script to compare student grades between two Excel snapshots.

# 📈 RoseDale Grade Comparison Tool

This is a self-contained Python script designed to compare student course grades between two chronological Excel snapshot files (e.g., Midterm vs. Final) for the same course. It calculates the **Grade Change** and identifies the **Most Improved Student(s)**.

---

## 🚀 1. Deployment and Quick Start (For End-Users)

The fastest way to use this tool is via the pre-built Windows executable.

1.  **Download:** Obtain the latest distribution folder containing the `dist` folder and the `.bat` file.
2.  **Data Placement:** Place your two comparison Excel files (e.g., `SHEN-MTH1Wa_grades_08Nov2025.xlsx`) directly into the root of the tool folder (next to the `.bat` file).
3.  **Run:** Double-click the **`Run_Grade_Comparison_Tool.bat`** file.
4.  Follow the on-screen prompts (Option 1 for defaults, Option 2 for manual file entry).
5.  **Output:** The report will be generated in the same folder, named like `[CourseName]_Grade_Comparison_Report_[Date].xlsx`.

---

## 🛠️ 2. Setting Up the Development Environment

If you need to make code changes (`grade_comparator.py`) or rebuild the executable, follow these steps.

### 2.1. Clone the Repository

1.  **Install Git:** Download and install Git from [git-scm.com](https://git-scm.com/).
2.  **Open Terminal/Command Prompt** and navigate to your desired parent directory.
3.  **Clone the project:**
    ```bash
    git clone [https://github.com/kpm25/RoseDale_Grade_Comparator.git](https://github.com/kpm25/RoseDale_Grade_Comparator.git)
    cd RoseDale_Grade_Comparator
    ```

### 2.2. Install Python and Dependencies

This project requires **Python 3.x**.

1.  **Verify Python:** Check if Python is installed (`python --version` or `python3 --version`). If not, install the latest version, ensuring you **Add Python to PATH** during installation.

2.  **Create and Activate a Virtual Environment** (Highly Recommended):
    ```bash
    python -m venv .venv
    # Activate on Windows (PowerShell/Git Bash):
    source .venv/Scripts/activate
    ```

3.  **Install Required Packages:**
    ```bash
    pip install -r requirements.txt
    ```

---

## 📦 3. Building the Standalone Executable

Once your environment is active, you can rebuild the distribution files.

### 3.1. Rebuild the Executable (`.exe`)

Run the following command while your virtual environment is active:



pyinstaller --onefile --console --name "Grade_Comparator_Tool" grade_comparator.py


### 3.2. Deployment Files

The new executable will be found in the **`dist`** folder. The complete deployment kit includes:
---
The Run_Grade_Comparison_Tool.bat contents:

@echo off
REM --- SETS THE CURRENT WORKING DIRECTORY TO WHERE THE BATCH FILE IS LOCATED ---
cd /d "%~dp0"

echo Running Grade Comparison Tool...
"dist\Grade_Comparator_Tool.exe"

echo.
echo Press any key to close this window and check the output file.
pause > nul

Known Data Handling
The script includes several robust features to ensure accurate processing:

Grade Scaling: Automatically converts decimal grades (e.g., 0.85) to percentages to ensure accurate comparison.

Whitespace Cleanup: Strips leading/trailing spaces from column headers, preventing KeyError crashes.

Flexible Date Parsing: Accepts dates in file names with single or double digits (e.g., 2Dec2025 or 12Dec2025).

Chronological Check: Validates that the older file has an equal or lower number of graded assessments than the newer file.


If you need any assistance with the Git commands to commit this file to your repository, just let me know!
```bash