# crm-hierarchy-reconciliation
Python-based CRM hierarchy reconciliation pipeline that processes 15,150 synthetic facility records, resolves parent network affiliations, detects data-quality exceptions, and generates a multi-sheet Excel reporting workbook.

# CRM Hierarchy Reconciliation Pipeline

**High-Volume Data Ingestion | Entity Resolution | Data Quality | Excel Reporting**

## Overview

This project demonstrates an automated approach to processing high-volume CRM data and reconciling organizational hierarchies.

It was developed as a technical assessment project to address two operational challenges: slow Excel processing of large datasets and inconsistent parent-network affiliations across local healthcare and senior-care facilities.

The solution uses Python to process synthetic CRM records, map facilities to their parent national networks, identify data-quality issues, and generate a structured Excel workbook for operational reporting and manual review.

**All organization and facility names in the demonstration dataset are fictional. The project does not contain real patient or healthcare records.**

## Project Objectives

- Process more than 15,000 CRM facility records efficiently.
- Reconcile local facilities with their parent national partner networks.
- Identify missing, inconsistent, misspelled, and unmapped affiliations.
- Detect incomplete records, invalid postal codes, and potential duplicates.
- Generate an Excel workbook with a dedicated exception-review workflow.
- Provide a repeatable pipeline that can process updated CRM exports.

## Technology Stack

- **Python** — pipeline development and data processing
- **Pandas** — dataset ingestion, transformation, and analysis
- **OpenPyXL** — Excel workbook generation, formatting, tables, and charts
- **difflib** — approximate string matching for misspelled network names
- **Regular expressions** — data validation
- **Microsoft Excel** — operational reporting and exception review

## Dataset

The synthetic dataset contains approximately 15,150 facility records, including 15,000 generated records and 150 additional duplicate records.

Supporting reference datasets include:

- `CRM_Facilities.csv` — facility records and deliberately introduced inconsistencies.
- `Partner_Networks.csv` — 10 fictional national partner networks.
- `Sub_Structures.csv` — 40 fictional branches linked to their parent networks.
- `Alias_Map.csv` — recognized alternative names and abbreviations.
- `Injected_Errors_Log.csv` — a record of deliberately introduced data-quality issues for validation.

The data generator uses a fixed random seed to support reproducible test data generation.

## How the Pipeline Works

### 1. Data ingestion and normalization

Loads the source CSV files, standardizes text values, and prepares the records and reference tables for matching.

### 2. Reference lookup construction

Builds in-memory lookup dictionaries for national networks, aliases, branches, and postal codes. This reduces repeated lookup work during reconciliation.

### 3. Hierarchy reconciliation

Resolves each facility's parent network using the following matching strategy:

1. Match the facility's branch to its known parent network.
2. Match the stated parent network against official names.
3. Resolve recognized aliases and abbreviations.
4. Detect branch names entered incorrectly in the parent field.
5. Apply fuzzy matching to eligible misspelled network names.
6. Use postal-code reference information when other mapping methods cannot resolve the network.

Each resolved record includes its reconciliation method and confidence level.

### 4. Data-quality and exception detection

Flags missing parent affiliations, unrecognized network names, conflicting hierarchies, unmapped branches, invalid or missing postal codes, missing agreement statuses, and possible duplicate records.

Exceptions are assigned severity levels to help operational teams prioritize their reviews. Suggested corrective actions are included where possible.

### 5. Excel workbook generation

Creates a formatted Excel workbook containing operational reporting and reconciliation outputs.

## Excel Workbook Structure

The generated `CRM_Reconciliation_Master.xlsx` is designed around the following worksheets:

| Worksheet | Purpose |
|---|---|
| Start Here | User instructions and severity definitions |
| Dashboard | Summary indicators, exception charts, and network-level coverage |
| Master Reconciled | Complete facility dataset with resolved affiliations |
| Exception Review | Records requiring inspection, suggested fixes, and reviewer fields |
| Partner Reference | National networks and their branches |
| Alias Map | Recognized alternative network names |
| QA Validation | Comparison of planted errors against detected issues |
| Pipeline Notes | Matching logic and performance design decisions |

## Performance and Design Decisions

The reconciliation logic uses in-memory lookup dictionaries rather than performing repeated row-level Excel lookups.

The workbook stores the processed reconciliation results as values and uses native Excel tables for filtering and sorting. Summary formulas and charts support operational monitoring without requiring extensive formula calculations across every record.

The pipeline separates data generation, reconciliation logic, and workbook creation so that the workflow can be maintained and extended.

## Quality Assurance

The data generator deliberately introduces known inconsistencies, and the error log provides a reference for testing whether the pipeline identifies them.

The QA worksheet compares planted error categories against detected exceptions. Detection results should be reviewed and validated before the pipeline is used with production data.

## Project Structure

```text
crm-hierarchy-reconciliation/
├── README.md
├── requirements.txt
├── generate_dummy_data.py
├── build_workbook.py
├── CRM_Facilities.csv
├── Partner_Networks.csv
├── Sub_Structures.csv
├── Alias_Map.csv
├── Injected_Errors_Log.csv
└── CRM_Reconciliation_Master.xlsx
```

The generated CSV files and workbook may be excluded from version control if repository size, privacy, or distribution requirements make it preferable to generate them locally.

## How to Run

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/crm-hierarchy-reconciliation.git
cd crm-hierarchy-reconciliation
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Generate the synthetic datasets

```bash
python generate_dummy_data.py
```

### 4. Build the Excel workbook

```bash
python build_workbook.py
```

The generated workbook will be saved as `CRM_Reconciliation_Master.xlsx`.

## Potential Applications

This approach can be adapted for CRM data cleanup, partner onboarding, organizational hierarchy management, entity resolution, data migration, and operational reporting in healthcare, senior care, nonprofit, and other multi-branch organizations.

## Limitations

This is a demonstration built on synthetic data. Matching rules, reference tables, and validation procedures would need to be reviewed against the requirements of a real organization before production deployment.

The workbook is an output of the Python pipeline; its contents do not automatically refresh from changed CSV files unless the pipeline is rerun.

## Author

**Lydia Vamba**

Data Analytics | Python | Data Quality | Data Processing | Operational Reporting

GitHub: [github.com/lydiavamba](https://github.com/lydiavamba)
