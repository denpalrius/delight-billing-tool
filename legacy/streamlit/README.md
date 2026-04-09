# 🏥 Delight Billing Tool

A Streamlit-powered daily staffing analysis tool that processes per-person Excel billing sheets and generates a consolidated daily summary workbook.

## Overview

This tool is designed for healthcare staffing teams to streamline the review and billing reconciliation process. It ingests one or more per-person `.xls` / `.xlsx` files, extracts service hours per provider and individual, and produces a clean Excel summary with daily matrices, provider totals, and 24-hour cap calculations.

## Features

- 📂 **Multi-file upload** — Process multiple Excel billing files at once
- 🔍 **Automatic section detection** — Locates all `Date` header blocks below the `Time Zone:` row in each sheet
- 🧮 **Hours aggregation** — Converts session minutes to decimal hours and groups by date, service provider, and individual
- 📊 **Daily Matrix output** — Generates one block per date with:
  - Service provider rows with hours per individual
  - Provider-level row totals (SUM formula)
  - Individual-level column totals
  - 24-hour cap remaining per individual
- 📥 **One-click download** — Download the summary as a dated `.xlsx` file

## Tech Stack

| Layer | Library |
|-------|---------|
| UI | [Streamlit](https://streamlit.io) |
| Excel parsing | [pandas](https://pandas.pydata.org), [xlrd](https://xlrd.readthedocs.io), [openpyxl](https://openpyxl.readthedocs.io) |
| Excel generation | [openpyxl](https://openpyxl.readthedocs.io) |

## Getting Started

### Prerequisites

- Python 3.10+
- `pip` or a virtual environment manager

### Installation

```bash
# Clone the repository
git clone https://github.com/denpalrius/delight-billing-tool.git
cd delight-billing-tool

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate   # on Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Running the App

```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

## Usage

1. Launch the app with `streamlit run app.py`.
2. Click **Browse files** and upload one or more per-person Excel billing sheets (`.xls` or `.xlsx`).
3. The tool will parse each file, aggregate service hours, and automatically generate the summary.
4. Click **Download Summary Excel** to save the `daily_summary_<date>.xlsx` file.

## Expected Input Format

Each uploaded Excel file should follow this structure:

- **Cell D3** — Individual's name (comma-separated; first part is used)
- A **"Time Zone:"** label in column A somewhere above the data sections
- One or more data sections starting with a row where **column A = "Date"** and ending with a row where **column C = "Total"**
- Within each data section:
  - **Column E (index 4)** — Duration in minutes
  - **Column G (index 6)** — Service Provider name

## Project Structure

```
delight-billing-tool/
├── app.py                        # Main Streamlit application
├── daily_staffing_analysis.ipynb # Jupyter notebook for exploratory analysis
├── requirements.txt              # Python dependencies
└── README.md                     # Project documentation
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

This project is proprietary. All rights reserved.
