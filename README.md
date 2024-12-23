# harmonize_joe_ejm
# Combine and Harmonize JOE and EJM Listings

This repository contains a Python script that merges and harmonizes data from two job-ad listing platforms for economists:

- **JOE** (Job Openings for Economists, by the AEA)  
- **EJM** (Economics Job Market)

Both websites provide a download option for all listings. Once downloaded, place them in the correct folders (`joe_listings` for JOE, `ejm_listings` for EJM), ensure the filenames follow the expected naming pattern, and run the script. It reads the latest files, merges and harmonizes their structure, and appends them to (or creates) an Excel file called `Application_MasterFile.xlsx`.

> **Disclaimer**: This script is provided “as-is” for personal use. It is not guaranteed to be error-free or fit for any specific purpose. Feel free to adjust and modify as you see fit.

---

## Table of Contents

1. [Overview](#overview)  
2. [How It Works](#how-it-works)  
3. [Required Setup](#required-setup)  
4. [Usage Instructions](#usage-instructions)  
5. [Script Logic](#script-logic)  
   - [Environment Setup](#environment-setup)  
   - [Reading & Harmonizing JOE Listings](#reading--harmonizing-joe-listings)  
   - [Appending to Excel Master File (JOE)](#appending-to-excel-master-file-joe)  
   - [Reading & Harmonizing EJM Listings](#reading--harmonizing-ejm-listings)  
   - [Appending to Excel Master File (EJM)](#appending-to-excel-master-file-ejm)  
6. [Handling Deletions](#handling-deletions)  
7. [Key Features](#key-features)  
8. [Customization Notes](#customization-notes)  
9. [Disclaimer](#disclaimer)  
10. [Improved Code Suggestions](#improved-code-suggestions)

---

## Overview

- **Goal**: Combine job ads from the AEA’s JOE platform and the EJM platform into a single Excel workbook with two sheets: `Listings` (accepted listings) and `Deleted` (excluded listings or those manually removed).  
- **Reasoning**: The two data sources have different fields, naming conventions, and some listings can appear in both. This script normalizes the columns, merges new rows, and preserves any manual changes (e.g., notes, bold formatting) in the Excel file.

---

## How It Works

1. **Manually download** each platform’s listings.
2. **Place** them into their corresponding folders, renaming them as follows:
   - JOE: `joe_listings/joe_resultset_dd_mm_yyyy.xlsx` (or `.xls`)
   - EJM: `ejm_listings/positions_dd_mm_yyyy.csv`
3. **Run** the Python script.  
4. The script:
   - Finds the **most recent** listing file by date (based on the filename).
   - Reads and **harmonizes** the columns (removing or renaming certain ones).
   - Appends new rows to either the `Listings` or `Deleted` sheet of `Application_MasterFile.xlsx`.
   - If a row is manually deleted from `Listings`, the script will move it to the `Deleted` sheet on subsequent runs.

---

## Required Setup

- **Python**: Python 3.7+ recommended.
- **Python packages used**:
  - `pandas`
  - `openpyxl`
  - `datetime` (standard library)
  - `re` (standard library)
  - `os` (standard library)

Install missing packages with:

```bash
pip install pandas openpyxl
"

---

## 4. Usage Instructions

1. **Clone or Download** this repository.
2. **Install** any missing Python libraries (`pandas`, `openpyxl`) in your environment.
3. **Download** the job listings from:
   - **JOE**: Save the file(s) in the `joe_listings/` folder.
   - **EJM**: Save the file(s) in the `ejm_listings/` folder.
4. **Make sure** your downloaded files match the naming patterns expected:
   - `joe_resultset_dd_mm_yyyy` for JOE (e.g. `joe_resultset_23_12_2024.xlsx`)
   - `positions_dd_mm_yyyy` for EJM (e.g. `positions_23_12_2024.csv`)
5. **Run** the script:

```
python combine_listings.py
"

6. After the first run, you will see an **`Application_MasterFile.xlsx`** with two sheets:
   - **Listings** – All accepted listings so far.
   - **Deleted** – All excluded or removed listings.
7. If you see any unwanted listings in the **Listings** sheet, **delete** those rows manually, **save** the Excel file, and then **re-run** the script to move them to the **Deleted** sheet automatically.
8. Going forward, continue to manually download updated listing files, place them in the correct folders, and **re-run** the script to update the master file with the new rows.

---

## 5. Detailed Script Logic

### Environment Setup

```
if os.environ.get('USER') == 'javad':
    os.chdir('/Users/javad/Dropbox/JM')
elif os.environ.get('USERNAME') == 'javad_s':
    os.chdir('C:/Non-Roaming/javad_s/Dropbox/JM')
else:
    print("working directory was not changed.")
"

- The script attempts to set the working directory based on environment variables.
- If neither variable matches, it leaves the working directory unchanged.

### Reading & Harmonizing JOE Listings

1. The script searches for files in **`joe_listings/`** named like `joe_resultset_dd_mm_yyyy`. It uses a regex to find the date.  
2. It picks the **latest** file (by date) and reads it using `pandas.read_excel()`.  
3. Certain columns are **removed** (e.g. `'joe_issue_ID', 'jp_section', 'jp_full_text', ...`).  
4. Columns are **renamed** for consistency:

```
df_new.columns = [col if col == 'jp_id' else col.replace('jp_', '') for col in df_new.columns]
df_new = df_new.rename(columns={'Application_deadline': 'deadline'})
"

5. The `deadline` column is converted to a `datetime.date`.  
6. A **JOE URL** column (`joe_url`) is created for easy navigation:

```
df_new['joe_url'] = df_new['jp_id'].apply(
    lambda jp_id: f'https://www.aeaweb.org/joe/listing.php?JOE_ID={jp_id}'
)
"

7. A helper function **`reorder_and_fill_columns`** ensures the final DataFrame has a fixed column order and fills missing columns with `None`.

### Appending to Excel Master File (JOE)

1. The script then checks if **`Application_MasterFile.xlsx`** exists.
2. If it **does not** exist, it creates it with two sheets: **Listings** and **Deleted**.
3. If it **does** exist, it loads the workbook and reads both **Listings** and **Deleted** into `pandas` DataFrames.
4. It collects a set of **existing IDs** from both sheets.
5. Only **new** (unique) JOE postings are considered for addition.
6. The script excludes listings based on:
   - If the country is in the `excluded_countries` list (e.g., `'china', 'taiwan', 'philippines'`, etc.).
   - Or if a listing with the **same batch date** is already in the file.

   Those get appended to **Deleted**. Otherwise, they go to **Listings**.
7. Finally, it **appends** new rows to the workbook **without overwriting** existing rows, by using `openpyxl` and a helper function like:

```
def append_df_to_ws(ws, df):
    if df.empty:
        return
    rows = dataframe_to_rows(df, index=False, header=False)
    for row in rows:
        ws.append(row)
"

8. The workbook is then **saved**.

### Reading & Harmonizing EJM Listings

1. The script does a **similar** process for EJM:
   - Looks in **`ejm_listings/`** for `positions_dd_mm_yyyy`.
   - Picks the latest file, reads it (likely CSV) using `pandas.read_csv()`.
   - Renames columns to a standardized set (`ejm_id`, `ejm_url`, `title`, etc.).
   - Converts `deadline` to a `datetime.date`.
   - Calls the same **`reorder_and_fill_columns`** to align columns.
2. Like JOE, it merges into the **Excel** file, checking for **existing IDs** in both **Listings** and **Deleted**, and uses the **`excluded_countries`** rule to decide if a posting should go to **Listings** or **Deleted**.

### Appending to Excel Master File (EJM)

1. The EJM chunk appends new EJM listings to the **same** Excel file **`Application_MasterFile.xlsx`** but also uses the **same** logic for “batch” detection and existing IDs.
2. Finally, the script **saves** the updated workbook.

---

## 6. Handling Deletions

- If you **manually delete** a listing from the **Listings** sheet in the Excel file, re-running the script **detects** that the unique ID is no longer present in **Listings**.
- The script will **add** that listing to the **Deleted** sheet (if it isn’t already there).
- If you run the script again or multiple times, that entry will remain in **Deleted** and will **not** be re-added to **Listings** if it appears again in a new data file.

---

## 7. Key Features

1. **Partial Overwrites**: Only **new rows** are appended; existing rows are kept as-is. This ensures that if you manually modify the Excel file (e.g., add notes in cells, change formatting, color rows), your changes are not lost.
2. **Consistent Column Order**: By using `reorder_and_fill_columns`, all final data is aligned in a single consistent format.
3. **Handling Exclusions**: The script automatically excludes postings whose country is in a predefined list (`excluded_countries`).
4. **Batch Logic**: The code uses a combination of date-based naming and batch identification to decide if a new listing belongs in **Listings** or **Deleted**.

---

## 8. Customization Notes

- The list `excluded_countries` can be **modified** to your preference:

```
excluded_countries = [
    'china', 'taiwan', 'philippines', 'hong', 'malawi', 'colombia', 'japan', 
    'brazil', 'chile', 'india', 'israel', 'argentina', 'mexico', 'saudi arabia', 
    'turkey', 'south africa', 'bangladesh', 'lebanon', 'uzbekistan', 'russia', 
    'korea', 'saudi'
]
"

- You can **comment out** or remove the environment checks for the working directory if they don’t apply to your local setup.
- If you **want to change** the default column order or remove any columns, you can adapt the `master_columns` list and the relevant rename steps in the script:

```
master_columns = [
    'institution', 'division', 'department', 'keywords', 'title', 'deadline', 
    'country', 'jp_id', 'ejm_id', 'joe_url', 'ejm_url', 'BatchDate', 'Source'
]
"

- If your data files have a **different naming structure**, you’ll need to adjust the **regular expressions** and filename matching logic.

---

## 9. Disclaimer

This script is shared for **personal** or **academic** use. We make **no warranty** that it is error-free or complete. You are advised to verify the results (e.g., check that the correct rows are added to **Listings** or **Deleted**) and to back up any important data before use.

**Enjoy merging your job listings, and best of luck with your applications!**
