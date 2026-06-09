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
---

**## 🚀 How It Works & Configuration**

search_routine = "Main version" # Or "Ramadan Routine"
search_batch = "44"
search_section = "D"
- **Step 3: Execute the Script** — Fire up your terminal environment and run the following command to generate your filtered lookup data:
```bash
python find_routine.py




