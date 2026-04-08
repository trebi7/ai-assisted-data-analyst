# AI-Assisted Data Analyst
### WASH Data Analysis — Côte d'Ivoire 2026

This project demonstrates a **human-guided, AI-assisted** approach to data analysis — where the analyst drives every decision and Claude AI executes the technical work under direction. The subject is a WASH (Water, Sanitation and Hygiene) survey covering 1,000 households across two departments in Côte d'Ivoire.

> **How it works:** The analyst defines the objectives, chooses the methods, reviews the outputs, and steers each step. Claude translates those instructions into code, visualizations, and written content — acting as an AI copilot, not an autonomous agent.

---

## Guided Workflow — Step by Step

### 1. Data Collection & Structuring
The analyst specified the survey variables, geographic coverage, and field categories required for a WASH assessment. Under those instructions, Claude generated a realistic, field-ready dataset (`wash_test_data.xlsx`) with 1,000 records — mirroring what a KoBoToolbox or ODK form would produce.

### 2. Data Cleaning & Preparation
The analyst identified the data quality issues to address and defined the cleaning rules. Claude then applied them in Python:
- Standardization of categorical values (source types, quality labels)
- Handling of missing and inconsistent entries
- Aggregation by department, village, and indicator type
- Preparation of pivot-ready structures for Excel and chart generation

### 3. Data Processing & Analysis
The analyst determined which indicators to compute and how to segment the data. Claude implemented the logic:
- Frequency distributions and cross-tabulations
- Departmental comparisons (Goh vs. Loh Djiboua)
- KPI calculations: access rate, treatment rate, open defecation rate, handwashing rate
- Identification of critical gaps and priority zones

### 4. Data Visualization
The analyst chose the chart types, the indicators to display, and the visual style. Claude produced all charts using `matplotlib` (`generate_charts.py`, `generate_charts_en.py`):
- Donut chart: water supply sources breakdown
- Grouped bar chart: water quality by department
- Stacked bar chart: sanitation infrastructure by department
- Bar charts: hygiene practices, water treatment methods, gender-disaggregated data

All charts follow the official WASH color palette.

### 5. Report Generation
The analyst defined the report structure, sections, and narrative tone. Claude generated two automated report formats:
- **HTML Report** (`generate_html_report.py` / `_en.py`): responsive web report with embedded charts, KPI cards, key findings, and recommendations
- **Word Report** (`generate_word_report.py` / `_en.py`): formatted `.docx` with tables, charts, and narrative sections ready for institutional submission

All reports are available in both **French** and **English**.

### 6. Excel Dashboard
The analyst specified the dashboard layout, KPIs to highlight, and filtering dimensions. Claude automated the build using `win32com` and `openpyxl`:
- Raw data sheet with full dataset
- Summary sheets with KPIs and aggregated tables
- 11 native Excel charts (bar, pie, column, donut)
- 5 slicers: Department, Sub-prefecture, Water Source, Water Quality, Month
- Conditional formatting and branded color scheme

---

## Project Outputs

| File | Description |
|---|---|
| `wash_test_data.xlsx` | Source dataset (1,000 survey records) |
| `WASH_Dashboard.xlsx` | Excel dashboard — French |
| `WASH_Dashboard_EN.xlsx` | Excel dashboard — English |
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
| **Claude AI** | Code generation, visualization, report writing — under analyst direction |

---

## Key Findings (Summary)

- **Water access**: 49% of households rely on traditional or improved wells; only 8% are connected to the national utility (SODECI)
- **Water quality**: Over 40% of water sources are rated poor quality across both departments
- **Sanitation**: Less than 30% of households have access to improved latrines; open defecation remains prevalent
- **Hygiene**: Only 35% of households practice handwashing with soap at critical moments
- **Priority zones**: Rural villages in Loh Djiboua require urgent intervention on water quality and sanitation

---

## About This Project

This project illustrates how an analyst can leverage AI as a productivity tool — not a replacement for analytical thinking. The domain expertise, the methodological choices, and the quality control are entirely human. Claude AI handles the implementation: writing the code, generating the charts, formatting the reports. The result is a complete, professional-grade data analysis delivered at a fraction of the usual time.

---

*Built with Claude Sonnet 4.6 — Anthropic*
