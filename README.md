# Research Data Processing Scripts (SLR Workflow Support)

## Overview
This repository contains Python scripts developed for academic research workflows, with a focus on systematic literature review (SLR) data processing and journal quality classification.

The goal of sharing this repository is to provide **computational transparency** and a **reproducible workflow** for key data-handling steps (e.g., merging records, deduplication, and year-specific journal quartile matching).

## What’s inside
This repository includes scripts that support:

### 1) Bibliographic data processing
- Merging records exported from multiple sources (e.g., RIS/CSV exports)
- Cleaning and standardizing bibliographic metadata
- Deduplication using identifiers and metadata (e.g., DOI, title-based matching)

### 2) Journal ranking / quartile matching (SCImago SJR)
- Year-based matching of journals to **SJR quartiles (Q1–Q4)**
- Standardization of journal names to improve match accuracy
- Export of consolidated quartile outputs for downstream screening/ranking tasks

### 3) Research dataset preparation
- Producing structured outputs ready for screening/filtering and analysis
- Generating sorted and documented Excel outputs where relevant

## Key script: Year-Based Journal Quartile Matcher (1999–2024)
A core component of this repository is the **Year-Based Journal Quartile Matcher**, which:

- Reads SJR files across multiple subject areas (e.g., Computer Science, Psychology, Business) for **1999–2024**
- Extracts journal names and quartiles (Q1–Q4) (and rank where available)
- Merges sources into a unified dataset and standardizes journal names to identify duplicates correctly
- Matches each journal to its quartile **strictly by publication year**  
  (e.g., a row with year = 2007 is matched only using `SJR2007_QRank.xlsx`)
- Writes the matched quartile into the main Excel file (e.g., `Quartile_Matched`) and updates the relevant sheet

## Data availability
This repository does **not** include:
- Raw database exports (RIS/CSV)
- Extracted datasets
- Any proprietary or licensed data files

This is intentional due to:
- Academic database licensing restrictions
- Research integrity considerations
- Privacy, file size, and duplication constraints

The scripts are shared to demonstrate the workflow and enable reproduction **once equivalent input files are available**.

## Expected inputs (example)
Depending on the script, typical inputs may include:
- Exported bibliographic records (RIS/CSV)
- A “main” Excel file used for screening/ranking (e.g., `second filter.xlsx`)
- A folder containing year-specific SJR files (e.g., `1999–2023/SJR{year}_QRank.xlsx`)

> Note: File names and paths can be adjusted in the scripts to match your local structure.

## How to run (general)
1. Create a Python environment (recommended):
   - Python 3.9+
2. Install dependencies (as needed by scripts):
   - `pandas`, `openpyxl`, etc.
3. Place input files in the expected folders (or update file paths in the script).
4. Run the script:
   - `python year_based_journal_quartile_matcher.py`

## Reproducibility notes
To ensure consistent results:
- Keep input naming conventions stable (or update paths in the scripts)
- Use the same SJR source files and year definitions
- Review journal-name normalization rules used in matching

## Repository status
This repository may contain scripts created across different stages of a research project.  
Scripts are provided “as-is” for transparency and reproducibility support.

## Suggested citation
If you reference this repository in academic work, you may cite it as:
- *Author, Repository Title, GitHub repository, Year.*

(You can replace “Author” and “Year” as needed.)

## License
Unless otherwise stated, code is shared for academic and research use.  
(You may add an explicit license file if required by your publication.)

## Contact
If you need clarification on script usage or input structure, feel free to reach out (e.g., via GitHub issues or your preferred academic contact method).
