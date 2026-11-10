import os
import smtplib
import sys
from email.message import EmailMessage
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


def load_env_file(path: Path) -> dict[str, str]:
    values = {}
    if not path.exists():
        return values

    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip("\"'")
    return values


def setting(env: dict[str, str], name: str, default: str = "") -> str:
    return os.environ.get(name) or env.get(name, default)


def send_report(report_pdf: bytes, filename: str, env: dict[str, str]) -> str:
    host = setting(env, "SMTP_HOST")
    port = int(setting(env, "SMTP_PORT", "465"))
    username = setting(env, "SMTP_USER")
    password = setting(env, "SMTP_PASS")
    sender = setting(env, "MAIL_FROM", username)
    recipients = [
        address.strip()
        for address in setting(env, "KPI_EMAIL_RECIPIENTS").split(",")
        if address.strip()
    ]

    if not host or not sender or not recipients:
        raise RuntimeError("SMTP host, sender, and recipient must be configured.")
    if not username or not password:
        raise RuntimeError("SMTP username and password must be configured.")

    message = EmailMessage()
    message["Subject"] = "Monthly Management Summary Report"
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message.set_content(
        "Please find the monthly management summary report attached."
    )
    message.add_attachment(
        report_pdf,
        maintype="application",
        subtype="pdf",
        filename=filename,
    )

    with smtplib.SMTP_SSL(host, port, timeout=30) as server:
        server.login(username, password)
        refused = server.send_message(message)
    if refused:
        raise RuntimeError(f"SMTP refused recipients: {refused}")
    return ", ".join(recipients)


def main() -> None:
    filename = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "Monthly_Management_Summary_Report.pdf"
    )
    report_pdf = sys.stdin.buffer.read()
    if not report_pdf.startswith(b"%PDF-"):
        raise ValueError("No valid PDF was provided on standard input.")

    env = load_env_file(ROOT / ".env")
    recipients = send_report(report_pdf, filename, env)
    print(f"Report emailed directly without saving a copy: {filename}")
    print(f"Email accepted by SMTP for: {recipients}")


if __name__ == "__main__":
    main()
