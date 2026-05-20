from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "outputs"
ASSET_DIR = BASE_DIR / "assets"

INPUT_FILE = DATA_DIR / "panynj_workforce_metrics.xlsx"
REPORT_FILE = OUTPUT_DIR / "panynj_workforce_report.xlsx"
SUMMARY_FILE = OUTPUT_DIR / "executive_summary.txt"
ERROR_LOG_FILE = OUTPUT_DIR / "error_log.txt"

REQUIRED_COLUMNS = [
    "program_name",
    "participants",
    "completion_rate",
    "job_placement_rate",
    "funding_amount",
    "quarter",
]

OPTIONAL_COLUMNS = [
    "target_completion_rate",
    "target_job_placement_rate",
    "operating_cost",
    "sector_focus",
    "county",
    "demographic_focus",
    "logistics_specialization",
    "maritime_training_hours",
    "internship_pipeline_count",
    "employer_partnerships",
    "credential_earned_rate",
    "port_related_job_placements",
]

DOMAIN_SPECIALIZATIONS = [
    "Maritime Logistics",
    "Port Operations",
    "Freight Transportation",
    "Cargo Movement",
    "Warehousing",
    "Supply Chain Operations",
    "Transportation Infrastructure",
    "Distribution Networks",
    "Terminal Operations",
    "Trucking Logistics",
    "Rail Freight",
    "Intermodal Transportation",
    "Workforce Pipeline Development",
]

PANYNJ_COUNTIES = [
    "Bergen",
    "Essex",
    "Hudson",
    "Middlesex",
    "Monmouth",
    "New York",
    "Queens",
    "Richmond",
    "Union",
]

PRIMARY_COLOR = "#2f6f9f"
SECONDARY_COLOR = "#e9f2f8"
ACCENT_COLOR = "#2f80ed"
WARNING_COLOR = "#b7791f"
CRITICAL_COLOR = "#b91c1c"
