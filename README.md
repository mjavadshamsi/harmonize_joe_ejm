# harmonize_joe_ejm
# Combine and Harmonize JOE and EJM Listings

This repository contains a Python script that merges and harmonizes data from two job-ad listing platforms for economists:

- **JOE** (Job Openings for Economists, by the AEA)  
- **EJM** (Economics Job Market)

Both websites provide a download option for all listings. Once downloaded, place them in the correct folders (`joe_listings` for JOE, `ejm_listings` for EJM), ensure the filenames follow the expected naming pattern, and run the script. It reads the latest files, merges and harmonizes their structure, and appends them to (or creates) an Excel file called `Application_MasterFile.xlsx`.

> **Disclaimer**: This script is for personal/academic use. No claim is made that it is error-free or perfectly suited for all use cases.

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

- **Python 3.7+** (recommended)
- Install the following libraries if missing:
  ```bash
  pip install pandas openpyxl
