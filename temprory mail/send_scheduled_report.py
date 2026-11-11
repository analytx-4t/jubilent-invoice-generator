import importlib.util
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "reports" / "Monthly_Management_Summary_Report.pdf"
ATTACHMENT_NAME = "Invoice & Debit Note Generation Report.pdf"
MAILER_PATH = ROOT / "reports" / "monthly_management_summary_report.py"


def load_mailer():
    spec = importlib.util.spec_from_file_location("invoice_report_mailer", MAILER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load the existing mailer: {MAILER_PATH}")
    mailer = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mailer)
    return mailer


def main() -> None:
    if not REPORT_PATH.is_file():
        raise FileNotFoundError(f"Scheduled report PDF was not found: {REPORT_PATH}")

    report_pdf = REPORT_PATH.read_bytes()
    if not report_pdf.startswith(b"%PDF-"):
        raise ValueError(f"Scheduled report is not a valid PDF: {REPORT_PATH}")

    mailer = load_mailer()
    env = mailer.load_env_file(ROOT / ".env")
    recipients = mailer.send_report(report_pdf, ATTACHMENT_NAME, env)
    print(f"Scheduled report emailed successfully to: {recipients}")


if __name__ == "__main__":
    main()
