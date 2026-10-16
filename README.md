# Court List Generator — Python & Excel Automation

A Python project that automates the preparation of a formatted, next-day court-production list from spreadsheet records. The code illustrates data processing, Excel report generation, configurable sorting rules, conditional formatting, and print-ready layout.

**Portfolio context:** This repository shares source code for demonstrating Python and spreadsheet-automation skills. It does **not** contain operational data, official workbook templates, or real generated reports. Publication of the code should not be interpreted as endorsement or approval by an employer or government agency.

## What the program does

- Reads a `PrisonersList*.xls`, `.xlsx`, or `.xlsm` input file. The `.xls` reader supports an HTML-table export saved with an `.xls` extension; it is not a general-purpose legacy binary `.xls` reader.
- Extracts and transforms fields needed for a court list.
- Applies predefined court-area and within-court sorting, with names sorted A–Z within courts.
- Assigns serial numbers, with configurable handling of repeated JIDs within a court.
- Adds duplicate-value conditional formatting for selected columns.
- Generates an Excel report using a compatible template, including borders and landscape print setup.
- Defaults to a next-day report date, with a configurable date-prompt mode.

## Technologies

Python 3.10+ (recommended), `openpyxl`, and Python standard-library modules including `tkinter`, `datetime`, and `pathlib`.

## Installation

1. Install Python 3.10+ on Windows. Ensure `tkinter` is included in your Python installation.
2. Download or clone this repository.
3. In a terminal opened in the project folder, run:

   ```powershell
   python -m pip install -r requirements.txt
   ```

## Running the script

The script is designed for a **locally supplied, compatible** `COURT_TEMPLATE.xlsx` workbook and a matching `PrisonersList*.xls`, `.xlsx`, or `.xlsm` source file. Neither is included in this public repository.

1. Place your **authorised, non-sensitive** template named `COURT_TEMPLATE.xlsx` beside `court_list_generator.py`.
2. Place a compatible **fictional/test** input file beside the script, or use the file picker when prompted.
3. Run:

   ```powershell
   python court_list_generator.py
   ```

4. The script creates a dated Excel workbook in the project folder (the default date is tomorrow).

**Note:** Because the original template and source files are intentionally excluded, this repository is a code portfolio rather than a self-contained runnable demonstration. The separate fictional-data demo is better for people who want to run an example immediately.

## Configuration

Edit the user settings near the beginning of `court_list_generator.py`:

| Setting | Purpose |
| --- | --- |
| `TEMPLATE_NAME` | Name of the compatible local workbook template |
| `DATE_MODE` | Use tomorrow's date or ask for a date |
| `PAGE_ORIENTATION` | Excel page orientation |
| `SORT_BY` | Sorting configuration |
| `SAME_JID_SAME_SL_WITHIN_COURT` | Reuse serial numbers for repeated JIDs in the same court |
| `ADD_DUPLICATE_CONDITIONAL_FORMATTING` | Highlight duplicate values in selected columns |

## Repository contents

```text
court-list-generator-github/
├── court_list_generator.py
├── README.md
├── requirements.txt
└── .gitignore
```

## Data protection

- Do not commit or upload real prisoner or custody records, case details, personal data, or confidential materials.
- Do not upload official Excel templates or operational output reports without separate authorisation.
- Review the source code and comments for any work-specific information before publishing.
- The `.gitignore` file excludes common data formats but is not a substitute for reviewing the files selected for upload.

## Suggested résumé description

**Court List Generator | Python, Excel, openpyxl** — Developed a Python workflow for transforming spreadsheet data into structured, formatted court-list reports, with custom sorting, duplicate detection, serial numbering, and print-ready Excel output.
