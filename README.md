# DIU_Routine_Scrapper
# 🗓️ Smart Excel-to-JSON Routine Parser

A robust, flexible Python utility designed to parse complex, multi-sheet university class routines from Excel workbooks (`.xlsx`) and format them into clean, highly-structured JSON data. 

It includes advanced lookup filtering capabilities, allowing users to isolate schedules by **Routine/Sheet Type** (e.g., Main Version vs. Ramadan Routine), **Batch**, and **Section**.

---

## ✨ Features

- **Multi-Sheet Support:** Seamlessly reads and processes multiple routine grids from a single workbook file.
- **Automated Metadata Extraction:** Automatically parses and structures tags like `Semester`, `Department`, and `Effective From` directly from the spreadsheet headers.
- **Smart Data Structuring:** Cleans string data, standardizes messy Excel time dashes, and breaks complex cell values (like `SE216-43-E`) into structured keys: `courseCode`, `batch`, and `section`.
- **Granular Filtering:** Includes a built-in search utility to immediately isolate specific routines—perfect for student apps or quick lookup schedules.

---

## 📦 Tech Stack & Dependencies

- **Language:** Python 3.x
- **Libraries:** - `pandas` (Data processing & structure manipulation)
  - `openpyxl` (Modern Excel layout reader engine)
  - `json` (Built-in standard JSON serialization)

To install dependencies, run:
```bash
pip install pandas openpyxl

---

## ⚙️ Configuration & Code Changes

### 🔧 Step 1: Update the Script Variables
Open your `find_routine.py` file, scroll down to the bottom execution block, and change the values inside these quotes to match your class details:

- **Target Routine:** `search_routine = "Main version"` — *Change this to `"Ramadan Routine"` if needed.*
- **Target Batch:** `search_batch = "44"` — *Change this to your numeric batch identifier.*
- **Target Section:** `search_section = "D"` — *Change this to your specific section letter.*

### 🚀 Step 2: Run the Parser Command
Open your terminal window and type the execution command below to generate your personal JSON data file:

```bash
python find_routine.py

### 🚀 How It Works
search_routine = "Main version" # Or "Ramadan Routine"
search_batch = "44"
search_section = "D"
python find_routine.py
