# AI-Assisted Data Analyst Portfolio
### WASH Data Analysis — Côte d'Ivoire 2026

This project demonstrates how **Claude AI** (Anthropic) was used end-to-end as a data analyst assistant — from data collection to dashboard delivery — on a real-world WASH (Water, Sanitation and Hygiene) dataset covering two departments in Côte d'Ivoire.

---

## How Claude AI Powered the Entire Workflow

### 1. Data Collection & Structuring
Claude helped design the data collection structure by generating a realistic, field-ready dataset (`wash_test_data.xlsx`) with 1,000 survey records across variables such as water sources, water quality, sanitation infrastructure, hygiene practices, and geographic segmentation by department and village. This mirrors what a field data collection form (KoBoToolbox, ODK) would produce.

### 2. Data Cleaning & Preparation
Claude identified and applied data cleaning logic directly in Python:
- Standardization of categorical values (e.g., source types, quality labels)
- Handling of missing and inconsistent entries
- Aggregation by department, village, and indicator type
- Preparation of pivot-ready structures for Excel and chart generation

### 3. Data Processing & Analysis
Claude wrote all the data processing code, performing:
- Frequency distributions and cross-tabulations
- Departmental comparisons (Goh vs. Loh Djiboua)
- KPI calculations: access rate, treatment rate, open defecation rate, handwashing rate
- Identification of critical gaps and priority zones

### 4. Data Visualization
Claude generated all charts using `matplotlib` (`generate_charts.py`, `generate_charts_en.py`):
- Donut chart: water supply sources breakdown
- Grouped bar chart: water quality by department
- Stacked bar chart: sanitation infrastructure by department
- Bar charts: hygiene practices, water treatment methods, gender-disaggregated data

All charts follow the official WASH color palette and are exported for embedding in both HTML and Word reports.

### 5. Report Generation
Claude built two automated report generators:
- **HTML Report** (`generate_html_report.py` / `_en.py`): a fully styled, responsive web report with embedded charts, KPI cards, key findings, and recommendations — ready to share with stakeholders
- **Word Report** (`generate_word_report.py` / `_en.py`): a professional `.docx` report with formatted tables, inserted charts, and narrative analysis sections — ready for institutional submission

All reports are available in both **French** and **English**.

### 6. Excel Dashboard
Claude automated the creation of a professional Excel dashboard (`create_dashboard.py` / `_en.py`) using `win32com` and `openpyxl`, including:
- Raw data sheet with full dataset
- Summary sheets with KPIs and aggregated tables
- Native Excel charts (bar, pie, column)
- Conditional formatting and branded color scheme
- Structured layout with headers, borders, and print-ready formatting

---

## Project Outputs

| File | Description |
|---|---|
| `wash_test_data.xlsx` | Source dataset (1,000 survey records) |
| `WASH_Dashboard.xlsx` | Full Excel dashboard with charts and KPIs |
| `WASH_Rapport.html` | HTML report — French |
| `WASH_Report_EN.html` | HTML report — English |
| `WASH_Rapport.docx` | Word report — French |
| `WASH_Report_EN.docx` | Word report — English |

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3 | All scripting and automation |
| `openpyxl` | Excel file creation and formatting |
| `win32com` | Excel COM automation (charts, slicers) |
| `matplotlib` | Chart generation |
| `python-docx` | Word report generation |
| HTML / CSS | Responsive web report |
| **Claude AI** | Data design, cleaning logic, full code generation, analysis, narrative writing |

---

## Key Findings (Summary)

- **Water access**: 49% of households rely on traditional or improved wells; only 8% are connected to the national utility (SODECI)
- **Water quality**: Over 40% of water sources are rated poor quality across both departments
- **Sanitation**: Less than 30% of households have access to improved latrines; open defecation remains prevalent
- **Hygiene**: Only 35% of households practice handwashing with soap at critical moments
- **Priority zones**: Rural villages in Loh Djiboua require urgent intervention on water quality and sanitation

---

## About This Portfolio

This project was built as a **portfolio piece** to demonstrate AI-assisted data analysis capabilities for NGO, humanitarian, and development sector clients on Upwork. Every line of code, every chart, every report section, and every insight was produced in collaboration with **Claude (Anthropic)** — showcasing how AI can dramatically accelerate the full data analysis pipeline without sacrificing rigor or quality.

---

*Built with Claude Sonnet 4.6 — Anthropic*
