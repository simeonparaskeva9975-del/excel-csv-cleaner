# Excel/CSV Cleaner + Report Generator

A small Python desktop application for analyzing messy spreadsheets, applying conservative cleaning rules, and exporting a client-ready Excel report.

## Current status

Stage 1 is complete:

- Git-friendly project structure
- Shared analysis and cleaning data models
- Deterministic synthetic sales-data generator
- Messy CSV and Excel demonstration files
- A documented expectation file describing every intentional defect
- Automated tests for dataset dimensions, duplicates, blanks, missing values, case variants, numeric problems, and date problems

The GUI is intentionally deferred. Its layout and appearance will be agreed before implementation.

## Project structure

```text
excel-csv-cleaner/
├── main.py
├── cleaner_app/
├── scripts/
├── sample_data/
├── output/
├── screenshots/
├── tests/
├── requirements.txt
└── README.md
```

## Set up on Windows

```powershell
py -3.12 -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## Rebuild the synthetic data

```powershell
python scripts\generate_demo_data.py
```

The Python generator rebuilds the CSV and the test expectations. The included Excel version contains the same raw values.

## Run the tests

```powershell
python -m unittest discover -s tests -v
```

## Run the current entry point

```powershell
python main.py
```

## Cleaning policy

The application will analyze before changing anything. Exact duplicates, fully blank structures, and outer whitespace can be proposed as safe fixes. Missing-value replacement, ambiguous dates, case-sensitive identifiers, outliers, and uncertain type conversions will require review instead of being silently changed.

## Roadmap

1. File loading and non-destructive analysis
2. Conservative cleaning engine and detailed log
3. Excel report export
4. User-approved GUI design and implementation
5. End-to-end testing, screenshots, and Upwork portfolio copy

## License

This portfolio project is provided under the MIT License.
