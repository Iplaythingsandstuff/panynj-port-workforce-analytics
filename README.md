# PANYNJ Port Workforce Analytics Platform

Enterprise-style workforce development analytics dashboard for the Port Authority of New York & New Jersey Port Department, designed for the Port Policy & Planning Unit within the Business Solutions Division.

The application supports operational and strategic analysis of maritime logistics, cargo operations, freight movement, terminal operations, warehousing, trucking logistics, rail freight, intermodal transportation, and regional Transportation, Logistics, and Distribution workforce pipeline initiatives.

## Platform Overview

This Streamlit platform is built as an internal analytics system for workforce development reporting and executive decision support. It helps Port Department leadership evaluate program outcomes, monitor workforce readiness, compare funding efficiency, and prepare public-sector style reporting for maritime and logistics workforce initiatives.

Core use cases include:

- Workforce development program performance tracking
- Transportation and logistics labor pipeline analysis
- Council on Port Performance workforce programming support
- Strategic initiative evaluation
- Public-sector workforce metrics reporting
- Funding allocation and efficiency review
- Maritime and freight workforce readiness analysis

## Technology Stack

- Python
- Streamlit
- pandas
- Plotly
- openpyxl
- numpy

The application runs entirely offline and does not require paid APIs.

## Installation

```bash
cd port-authority-workforce-analytics
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Start the Dashboard

```bash
streamlit run app.py
```

## Deploy to Streamlit Cloud

This repository is ready for Streamlit Community Cloud deployment.

Main file path:

```text
app.py
```

Deployment instructions are included in:

```text
STREAMLIT_DEPLOYMENT.md
```

The app loads the default workbook at:

```text
data/panynj_workforce_metrics.xlsx
```

## Expected Input Format

Required columns:

- `program_name`
- `participants`
- `completion_rate`
- `job_placement_rate`
- `funding_amount`
- `quarter`

Optional columns are automatically incorporated when present:

- `target_completion_rate`
- `target_job_placement_rate`
- `operating_cost`
- `sector_focus`
- `county`
- `demographic_focus`
- `logistics_specialization`
- `maritime_training_hours`
- `internship_pipeline_count`
- `employer_partnerships`
- `credential_earned_rate`
- `port_related_job_placements`

Quarter values should use the format `Q1 2025`, `Q2 2025`, and so on.

## Using Real Data

The dashboard includes a sidebar data source control.

Options:

- Use the built-in sample workbook
- Upload a real `.xlsx` workforce metrics workbook
- Download the Excel template from the sidebar

When a real workbook is uploaded, the dashboard reruns validation, KPIs, rankings, visualizations, executive summary text, and Excel export using the uploaded data.

## Dashboard Features

- Executive KPI cards for participants, completion, placement, funding, active programs, and maritime placement
- Sidebar filters for quarter, program, county, logistics specialization, participants, and readiness score
- Workforce participation trends by quarter
- Logistics-sector placement trends
- Top and bottom program ranking engine
- Funding allocation vs. placement outcome scatter analysis
- Logistics specialization participant distribution
- Workforce readiness heatmap
- Completion vs. placement stacked bar analysis
- County-level workforce participation summary
- Automated executive summary generation
- Excel reporting export with formatted sheets and conditional formatting
- Download buttons for Excel report, executive summary, and cleaned dataset

## Workforce Readiness Score

The readiness score is normalized to a 0-100 scale using:

```text
completion_rate * 0.30
+ job_placement_rate * 0.35
+ credential_earned_rate * 0.15
+ port_related_job_placements score * 0.10
+ participant_growth_score * 0.10
```

## Repository Structure

```text
port-authority-workforce-analytics/
├── app.py
├── analytics.py
├── config.py
├── requirements.txt
├── README.md
├── .gitignore
├── data/
│   └── panynj_workforce_metrics.xlsx
├── outputs/
│   ├── panynj_workforce_report.xlsx
│   ├── executive_summary.txt
│   └── error_log.txt
├── assets/
│   ├── panynj_logo_placeholder.png
│   └── dashboard_banner.png
└── docs/
    └── sample_dashboard_screenshots.md
```

## Screenshots

Dashboard screenshots can be added to `docs/sample_dashboard_screenshots.md` after running the application locally.

Suggested screenshots:

- Executive KPI overview
- Workforce participation and placement trend analysis
- Program ranking and workforce readiness view
- Funding efficiency analysis
- County and logistics specialization summaries

## Outputs

The application generates:

- `outputs/panynj_workforce_report.xlsx`
- `outputs/executive_summary.txt`
- `outputs/error_log.txt`

Excel report sheets:

1. Cleaned Workforce Data
2. KPI Summary
3. Trend Analysis
4. Top Programs
5. Bottom Programs
6. Strategic Insights
7. Funding Analysis
8. Workforce Readiness Scores

## Future Improvement Roadmap

- Add benchmark targets tied to annual Port Department strategic goals
- Add historical year-over-year workforce cohort analysis
- Add employer partner and training provider scorecards
- Add automated PDF briefing export
- Add scenario modeling for funding changes and placement targets
- Add geospatial mapping when approved internal data is available
- Add role-based views for executive, analyst, and program manager users

## GitHub Setup

Suggested repository name:

```text
panynj-port-workforce-analytics
```

Initialize and commit:

```bash
git init
git add .
git commit -m "Initial PANYNJ port workforce analytics platform"
```

Push to GitHub:

```bash
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/panynj-port-workforce-analytics.git
git push -u origin main
```
