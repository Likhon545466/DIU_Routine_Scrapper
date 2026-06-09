# DIU Routine Scrapper

## 🗓️ Smart Excel-to-JSON Routine Parser

A robust, flexible Python utility designed to parse complex, multi-sheet university class routines from Excel workbooks (`.xlsx`) and format them into clean, highly-structured JSON data.

It includes advanced lookup filtering capabilities, allowing users to isolate schedules by Routine/Sheet Type, Batch, and Section.

---

## ✨ Features

* Multi-Sheet Support. Reads and processes multiple routine grids from a single workbook file.
* Automated Metadata Extraction. Parses and structures tags like Semester, Department, and Effective From directly from spreadsheet headers.
* Smart Data Structuring. Cleans string data, standardizes messy Excel time dashes, and breaks complex cell values like `SE216-44-D` into structured keys:

  * `courseCode`
  * `batch`
  * `section`
  * `Teacher's Initial`
* Granular Filtering. Includes a built-in search utility to isolate specific routines instantly.

---

## 📦 Tech Stack & Dependencies

Language:

* Python 3.x

Libraries:

* `pandas`
* `openpyxl`
* `json`

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## ⚙️ Configuration & Code Changes

### 🔧 Step 1: Modify Script Variables

Open `Scrapper_Code.py` and update the filter values:

```python


search_routine = "Main version"   # Example: "Ramadan Routine"
search_batch = "44"
search_section = "D"
```

---

## 🚀 Step 2: Run the Parser Command

Open your terminal and run:

```bash
python Scrapper_Code.py
```

---

## 🚀 How It Works

Example configuration:

```python id="h6v73n"
search_routine = "Main version"
search_batch = "44"
search_section = "D"
```

Run the script:

```bash id="ec6rrx"
python Scrapper_Code.py
```

The parser will:

* Read the Excel workbook
* Process all routine sheets
* Extract metadata
* Filter routines based on your configuration
* Generate clean JSON output
