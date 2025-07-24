# Nessus CSV to Consolidated Vulnerability Report Converter By Srikqr

This project contains a Python script (`tool_report_Separate.py`) that transforms one or more Nessus CSV export files into a single, well-structured Excel workbook. The resulting workbook makes it easier to track, remediate, and report vulnerabilities discovered during network scans.

## Key Features

* 🛠 **Auto-installs dependencies** – Missing Python packages (pandas, openpyxl, xlsxwriter, numpy) are automatically detected and installed.
* 📑 **Data cleansing & enrichment** – Filters out “None” risk items, de-duplicates findings, and aggregates host/protocol/port combinations per vulnerability.
* 📊 **Excel output with multiple tabs** – Produces a fully-formatted workbook with:
  * **Full Report** – All findings
  * **Critical / High / Medium / Low** – Separate sheets by risk rating (only if data exists)
* 🎨 **Professional formatting** – Wide columns, wrapped text, frozen header row, and auto-filters ready for management review.

## Prerequisites

* Python 3.8 or newer (tested on 3.8–3.12)
* Internet access during the first run so that missing packages can be installed automatically

> Tip: If your environment restricts outbound internet access, manually install the required libraries first:
>
> ```bash
> pip install pandas openpyxl xlsxwriter numpy
> ```

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/<your-org>/<your-repo>.git
   cd <your-repo>
   ```
2. (Optional) Create and activate a virtual environment.
3. You are ready to run the script – it will install any missing dependencies on first execution.

## Usage

```bash
python tool_report_Separate.py scan_results.csv
```

You can also process multiple files or use wildcards:

```bash
# Multiple explicit files
python tool_report_Separate.py week1.csv week2.csv

# All CSVs in the current directory
python tool_report_Separate.py *.csv
```

The script will:

1. Create a `final/` sub-directory (if it does not already exist).
2. For each input CSV, generate an Excel workbook named `<original>_final_report.xlsx` inside `final/`.

Each workbook contains the sheets described in *Key Features*.

## Column Mapping

The script is tolerant of column name variations commonly found in Nessus exports. It maps lowercase/uppercase equivalents to the following required fields:

| Expected Column | Purpose |
|-----------------|---------|
| Plugin ID | Unique identifier for the vulnerability plugin |
| CVSS Score | Numeric CVSS base score |
| Risk | Severity level (Critical/High/Medium/Low/None) |
| Host / Protocol / Port | Asset identifiers |
| Name | Vulnerability title |
| Synopsis / Description | Contextual details |
| Solution | Recommended fix |
| Plugin Output | Evidence from the scan |

## Example

After running:

```bash
python tool_report_Separate.py corp_scan.csv
```

you will find `final/corp_scan_final_report.xlsx` with a **Full Report** sheet and, for example, a **High** sheet summarizing all High-risk findings.

## Troubleshooting

* **No output generated?** Ensure your input is the Nessus *CSV* export, not HTML or XML.
* **Permission denied on package install?** Run inside a virtualenv or use `python -m pip install --user ...` to install packages locally.
* **Headers not recognized?** Confirm your CSV contains the columns listed in *Column Mapping*.

## Contributing

Pull requests are welcome! Please:

1. Fork the repo and create your branch: `git checkout -b feature/improve-mapping`.
2. Commit your changes: `git commit -m "Improve column mapping"`.
3. Push to the branch: `git push origin feature/improve-mapping`.
4. Open a pull request describing the enhancement or bug fix.

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Acknowledgements

* Nessus® is a registered trademark of Tenable™, Inc. This project is **not** affiliated with or endorsed by Tenable.
* Thanks to the open-source community behind `pandas`, `numpy`, and `xlsxwriter` for making data wrangling simple!
