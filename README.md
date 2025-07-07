# 🧾 Sales Co Apps 2.0 - Autocount Data Integrator

A **Streamlit** app to clean, merge, classify, and export **AutoCount Excel data** for use in Sales Co workflows, integrated with **Google Sheets** for model lookups.

[👉 **Access the Live App Here**](https://ga-sales-co-apps-data-integrator.streamlit.app/)

---

## 📑 Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Example Workflow](#example-workflow)
- [Project Structure](#project-structure)
- [Reference](#reference)

---

## ✨ Features

- **Excel File Upload**
  - Supports uploading 1 or 2 `.xlsx` / `.xls` files.
- **RTF Cleaning**
  - Automatically cleans Rich Text in descriptions.
- **Detail Extraction**
  - Extracts details and remarks from descriptions into structured columns.
- **Order Classification**
  - Labels orders as:
    - New Order
    - Warranty
    - Customade
    - Fixed/Removable/Inner Part
- **Google Sheets Lookup**
  - Matches item codes and models.
- **Interactive Reordering**
  - Drag & drop columns before export.
- **Export & Copy**
  - Download CSV or copy tab-delimited text for Sales Co Apps.

---

## ⚙️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-org/salesco-autocount-integrator.git

   cd salesco-autocount-integrator
   ```

2. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

3. **Create** `.streamlit/secrets.toml`

   ```toml
   [gsheets]
   credentials = "<YOUR GOOGLE SERVICE ACCOUNT JSON>"
   spreadsheet = "<YOUR GOOGLE SHEET URL>"
   ```

4. **Run the app**

   ```bash
   streamlit run app.py
   ```

---

## ⚙️ Configuration

* **Google Sheets Connection**

  * Your service account must have Editor access.
  * Sheet must contain:

    * `item code`
    * `model`
* **Column Mapping**

  * Certain columns are renamed for Sales Co output automatically.

---

## 🛠️ Usage

1. **Upload Files**

   * Upload one or two Excel files.
2. **Processing**

   * RTF cleaning and detail extraction applied.
   * Files merged (if 2 uploaded) by `PI`.
3. **Review & Filter**

   * Confirm extracted details, order types, and matched models.
4. **Column Selection**

   * Pick columns you need.
   * Drag & drop to reorder.
5. **Export**

   * Download CSV or copy TSV text.

---

## 💡 Example Workflow

**Scenario:**

* You received two AutoCount exports:

  * `sales_order1.xlsx`
  * `sales_order2.xlsx`
* You need to merge, clean, and prepare data for import.

**Steps:**

1. Open the app.
2. Upload both Excel files.
3. Review cleaned and merged data.
4. Select columns.
5. Reorder columns to match Sales Co import template.
6. Copy TSV text or download CSV.
7. Paste into Sales Co Apps.

**Daily Habit Example:**

* After each new batch of orders:

  * Upload files.
  * Check extraction.
  * Save/export data.
  * Keep Google Sheets lookup updated.

---

## 📂 Project Structure

```bash
salesco-autocount-integrator/
├── app.py                   # Main Streamlit app
├── requirements.txt         # Dependency list
└── README.md                # This README file
```

---

## 📚 Reference

* [Streamlit Documentation](https://docs.streamlit.io)
* [striprtf](https://github.com/caolan/striprtf)
* [streamlit\_gsheets](https://github.com/streamlit/streamlit-gsheets)
* [Google Sheets API](https://developers.google.com/sheets/api)
* [👉 Access the Live App](https://ga-sales-co-apps-data-integrator.streamlit.app/)
