# Internship Spreadsheet Generator

Scripts that scrape tech internship listings from GitHub repositories and output styled Excel spreadsheets.

## Data Sources

Internship data is pulled from the following community-maintained repositories:

- **[Canadian-Tech-Internships-2026](https://github.com/negarprh/Canadian-Tech-Internships-2026)** by [negarprh](https://github.com/negarprh) — Canadian tech internship listings in Markdown table format.
- **[Summer2026-Internships](https://github.com/SimplifyJobs/Summer2026-Internships)** by [SimplifyJobs](https://github.com/SimplifyJobs) — Comprehensive internship listings in HTML table format.

## Setup

Requires Python 3.12+.

```bash
pip install requests openpyxl
```

## Usage

### fetch_internships.py

Fetches listings from the Canadian-Tech-Internships repo and filters for **Western Canada** locations (BC, AB, SK, MB). Outputs a single-sheet Excel file with open listings sorted to the top.

```bash
python fetch_internships.py                    # output: western_canada_internships.xlsx
python fetch_internships.py -o custom.xlsx     # custom output path
python fetch_internships.py --open-only        # exclude closed listings
```

### fetch_simplify.py

Fetches listings from the SimplifyJobs repo and classifies them by region into separate sheets: Western Canada, Eastern Canada, Western US, and Other. Tracks changes between runs so new and closed listings are highlighted.

```bash
python fetch_simplify.py                          # output: simplify_internships.xlsx
python fetch_simplify.py -o custom.xlsx           # custom output path
python fetch_simplify.py --tracking-file t.json   # custom tracking file path
```

## Automation

A GitHub Actions workflow (`.github/workflows/update.yml`) runs `fetch_internships.py` weekly on Sundays at 8 AM UTC and auto-commits the updated spreadsheet.
