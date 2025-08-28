#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Email report generator for scheduled loads vs consumption.

Cross-platform:
- Windows: sends via Outlook (if available)
- Linux/macOS: sends via SMTP, or use --dry-run to save HTML locally

Usage example (Linux, dry run):
  python email_report.py \
    --carts-file \\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias\ \&\ Underbias_Trigger\Carts_Consumption.xlsx \
    --fmc-file \\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias\ \&\ Underbias_Trigger\fmc-data-source(5).xlsx \
    --date-start-serial 45894 \
    --days 7 \
    --to dhimv@amazon.com --cc test@amazon.com \
    --subject "Review Scheduled Loads as per Consumption" \
    --dry-run --output-html /workspace/report.html

Note: UNC paths (\\server\share\...) require the share to be mounted on Linux.
"""

from __future__ import annotations

import argparse
import datetime
import os
import platform
import sys
from dataclasses import dataclass
from typing import List, Optional


def _get_pandas():
    """Import pandas lazily to allow --skip-inputs dry-runs without it installed."""
    import importlib
    return importlib.import_module("pandas")


def excel_serial_to_date(serial_value: int) -> datetime.date:
    """Convert Excel 1900-date-system serial to date.

    Excel incorrectly treats 1900 as leap year. Using base 1899-12-30 handles this.
    """
    base = datetime.date(1899, 12, 30)
    return base + datetime.timedelta(days=int(serial_value))


def date_to_excel_serial(date_value: datetime.date) -> int:
    """Convert date to Excel 1900-date-system serial integer."""
    base = datetime.date(1899, 12, 30)
    return (date_value - base).days


def today_excel_serial() -> int:
    return date_to_excel_serial(datetime.date.today())


def coerce_planned_dock_arrival_to_serial(df, column_name: str = "Planned Dock Arrival"):
    """Return a Series with Excel-serial integers for the planned dock arrival column.

    Supports numeric serials or pandas datetimes; non-coercible values become NA.
    """
    pd = _get_pandas()
    series = df[column_name]
    # If already numeric, floor to int
    numeric = pd.to_numeric(series, errors="coerce")
    if numeric.notna().any() and numeric.dtype.kind in {"i", "u", "f"}:
        return numeric.apply(lambda x: int(x) if pd.notna(x) else pd.NA)

    # Try datetime-like
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.date.apply(lambda d: date_to_excel_serial(d) if pd.notna(d) else pd.NA)

    # Try parsing as datetimes
    parsed = pd.to_datetime(series, errors="coerce")
    if parsed.notna().any():
        return parsed.dt.date.apply(lambda d: date_to_excel_serial(d) if pd.notna(d) else pd.NA)

    # Fallback: all NA
    return pd.Series(pd.NA, index=df.index)


ALLOWED_SITE_TYPES = {"FC", "SC", "DS"}


@dataclass
class ReportRow:
    region: str
    site_type: str
    site: str
    date_str: str
    needed_tfr_80: float
    needed_tfr_93: float
    scheduled_trucks: int
    difference: float


def build_table(
    final_df,
    fmc_df,
    date_start_serial: int,
    num_days: int,
    current_serial_threshold: Optional[int] = None,
) -> "object":
    """Compute table similar to the provided logic.

    Assumes `final_df` has columns [Region, Type, Site] followed by date-derived columns:
    - Needed TFR 80 in columns index 10..(10+num_days-1)
    - Needed TFR 93 in columns index 17..(17+num_days-1)
    This matches K:Q and R:X (0-based iloc indexing).
    """
    pd = _get_pandas()

    if current_serial_threshold is None:
        current_serial_threshold = today_excel_serial()

    required_cols = {"Region", "Type", "Site"}
    missing = [c for c in required_cols if c not in final_df.columns]
    if missing:
        raise ValueError(f"Final_Sheet missing required columns: {missing}")

    # Ensure FMC has Stop and Action Type
    for c in ("Stop", "Action Type"):
        if c not in fmc_df.columns:
            raise ValueError(f"FMC data missing required column: {c}")

    # Coerce planned arrival to Excel serial int and add a floor-day Date column
    if "Planned Dock Arrival" not in fmc_df.columns:
        raise ValueError("FMC data missing required column: Planned Dock Arrival")

    fmc_serials = coerce_planned_dock_arrival_to_serial(fmc_df, "Planned Dock Arrival")
    fmc_df = fmc_df.copy()
    fmc_df["Date"] = fmc_serials

    date_serials: List[int] = [date_start_serial + i for i in range(num_days)]
    date_strs: List[str] = [excel_serial_to_date(d).strftime("%m/%d/%Y") for d in date_serials]

    table_rows: List[ReportRow] = []

    for _, row in final_df.iterrows():
        site_type = str(row["Type"]).strip()
        if site_type not in ALLOWED_SITE_TYPES:
            continue

        region = row["Region"]
        site = row["Site"]

        for i, serial_date in enumerate(date_serials):
            if serial_date < current_serial_threshold:
                continue

            needed80_col = 10 + i
            needed93_col = 17 + i

            # Bounds guard
            if needed80_col >= len(row.index) or needed93_col >= len(row.index):
                continue

            needed80_raw = row.iloc[needed80_col]
            needed93_raw = row.iloc[needed93_col]

            needed80 = float(needed80_raw) if pd.notna(needed80_raw) else 0.0
            needed93 = float(needed93_raw) if pd.notna(needed93_raw) else 0.0

            if site_type in {"FC", "SC"}:
                mask = (
                    (fmc_df["Stop"] == site)
                    & (fmc_df["Action Type"] == "DROPOFF")
                    & (fmc_df["Date"] == serial_date)
                )
            else:  # DS
                mask = (
                    (fmc_df["Stop"] == site)
                    & (fmc_df["Action Type"] == "PICKUP")
                    & (fmc_df["Date"] == serial_date)
                )

            count_have = int(mask.sum())
            diff = round(needed80 - count_have, 2)

            date_str = date_strs[i]

            table_rows.append(
                ReportRow(
                    region=str(region),
                    site_type=site_type,
                    site=str(site),
                    date_str=date_str,
                    needed_tfr_80=round(needed80, 2),
                    needed_tfr_93=round(needed93, 2),
                    scheduled_trucks=count_have,
                    difference=diff,
                )
            )

    df = pd.DataFrame([r.__dict__ for r in table_rows])
    if not df.empty:
        df = df[
            [
                "region",
                "site_type",
                "site",
                "date_str",
                "needed_tfr_80",
                "needed_tfr_93",
                "scheduled_trucks",
                "difference",
            ]
        ]
        df.columns = [
            "Region",
            "Type",
            "Site",
            "Date",
            "Needed TFR 80",
            "Needed TFR 93",
            "Scheduled Trucks",
            "Difference",
        ]
    return df


def dataframe_to_html_table(table_df) -> str:
    html_table = table_df.to_html(
        index=False,
        border=1,
        classes="table table-striped",
        formatters={
            "Needed TFR 80": lambda x: f"{x:.2f}",
            "Needed TFR 93": lambda x: f"{x:.2f}",
            "Difference": lambda x: f"{x:.2f}",
        },
        escape=False,
    )
    return html_table


def empty_table_html() -> str:
    headers = [
        "Region",
        "Type",
        "Site",
        "Date",
        "Needed TFR 80",
        "Needed TFR 93",
        "Scheduled Trucks",
        "Difference",
    ]
    thead = "".join(f"<th>{h}</th>" for h in headers)
    return f'<table class="table table-striped" border="1"><thead><tr>{thead}</tr></thead><tbody></tbody></table>'


def build_email_html(body_table_html: str) -> str:
    return f"""
<html>
<head>
<style>
.table {{
    border-collapse: collapse;
    width: 100%;
    font-family: Arial, sans-serif;
}}
.table th, .table td {{
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}}
.table th {{
    background-color: #f2f2f2;
    font-weight: bold;
}}
.table tr:nth-child(even) {{
    background-color: #f9f9f9;
}}
</style>
</head>
<body>
<p>Hi Team,</p>
<p>Please review the scheduled loads as per consumption.</p>
{body_table_html}
<p>Best Regards,</p>
<p>Vipin Dhiman</p>
</body>
</html>
""".strip()


def send_email_via_outlook(subject: str, html_body: str, to: str, cc: Optional[str] = None) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("win32com is not available on this platform") from exc

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    if cc:
        mail.CC = cc
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()


def send_email_via_smtp(
    subject: str,
    html_body: str,
    to: str,
    cc: Optional[str] = None,
    smtp_host: str = "",
    smtp_port: int = 587,
    smtp_user: Optional[str] = None,
    smtp_password: Optional[str] = None,
    smtp_use_starttls: bool = True,
    mail_from: Optional[str] = None,
) -> None:
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    recipients = [addr.strip() for addr in to.split(",") if addr.strip()]
    cc_list = [addr.strip() for addr in cc.split(",") if cc] if cc else []
    all_recipients = recipients + cc_list

    if not smtp_host:
        raise ValueError("SMTP host is required for SMTP sending")

    if not mail_from:
        mail_from = smtp_user or (recipients[0] if recipients else "noreply@example.com")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = mail_from
    msg["To"] = ", ".join(recipients)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        if smtp_use_starttls:
            server.starttls()
            server.ehlo()
        if smtp_user and smtp_password:
            server.login(smtp_user, smtp_password)
        server.sendmail(mail_from, all_recipients, msg.as_string())


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate and send scheduled loads vs consumption email report")

    parser.add_argument("--carts-file", required=False, help="Path to Carts_Consumption.xlsx (Final_Sheet)")
    parser.add_argument("--fmc-file", required=False, help="Path to fmc-data-source workbook")
    parser.add_argument("--fmc-sheet", default="fmc-data-source(5)", help="Sheet name in FMC workbook")

    parser.add_argument("--date-start-serial", type=int, required=True, help="Excel serial for first date column (e.g., 45894)")
    parser.add_argument("--days", type=int, default=7, help="Number of consecutive days (default: 7)")
    parser.add_argument(
        "--current-serial",
        type=int,
        default=None,
        help="Only include dates >= this Excel serial (default: today)",
    )

    parser.add_argument("--to", required=True, help="Comma-separated list of recipients")
    parser.add_argument("--cc", default=None, help="Comma-separated list of CC recipients")
    parser.add_argument("--subject", default="Review Scheduled Loads as per Consumption", help="Email subject")

    # Delivery options
    parser.add_argument("--use-outlook", action="store_true", help="Force Outlook sending (Windows only)")
    parser.add_argument("--use-smtp", action="store_true", help="Force SMTP sending")
    parser.add_argument("--dry-run", action="store_true", help="Do not send, just save HTML if --output-html provided")
    parser.add_argument("--output-html", default=None, help="Path to save generated HTML body (for dry run or debug)")
    parser.add_argument("--skip-inputs", action="store_true", help="Skip reading Excel inputs and generate empty table")

    # SMTP configuration (env overrides)
    parser.add_argument("--smtp-host", default=os.getenv("SMTP_HOST", ""), help="SMTP host")
    parser.add_argument("--smtp-port", type=int, default=int(os.getenv("SMTP_PORT", "587")), help="SMTP port")
    parser.add_argument("--smtp-user", default=os.getenv("SMTP_USER", None), help="SMTP username")
    parser.add_argument("--smtp-password", default=os.getenv("SMTP_PASS", None), help="SMTP password")
    parser.add_argument(
        "--smtp-starttls",
        dest="smtp_starttls",
        action="store_true",
        default=os.getenv("SMTP_STARTTLS", "true").lower() not in {"0", "false", "no"},
        help="Use STARTTLS (default: true)",
    )
    parser.add_argument("--from", dest="mail_from", default=os.getenv("SMTP_FROM", None), help="From address for SMTP")

    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    current_serial = args.current_serial if args.current_serial is not None else today_excel_serial()

    if args.skip_inputs:
        # Build HTML without requiring pandas
        html_table = empty_table_html()
    else:
        # Read Excel inputs
        if not args.carts_file or not args.fmc_file:
            print("--carts-file and --fmc-file are required unless --skip-inputs is used")
            return 2
        try:
            pd = _get_pandas()
            final_df = pd.read_excel(args.carts_file, sheet_name="Final_Sheet")
        except Exception as exc:
            print(f"Error reading Final_Sheet from {args.carts_file}: {exc}")
            return 2

        try:
            fmc_df = pd.read_excel(args.fmc_file, sheet_name=args.fmc_sheet)
        except Exception as exc:
            print(f"Error reading {args.fmc_sheet} from {args.fmc_file}: {exc}")
            return 2

        # Build table
        try:
            table_df = build_table(
                final_df=final_df,
                fmc_df=fmc_df,
                date_start_serial=args.date_start_serial,
                num_days=args.days,
                current_serial_threshold=current_serial,
            )
        except Exception as exc:
            print(f"Error building table: {exc}")
            return 3
        html_table = dataframe_to_html_table(table_df)

    # Generate HTML
    html_body = build_email_html(html_table)

    if args.output_html:
        try:
            with open(args.output_html, "w", encoding="utf-8") as f:
                f.write(html_body)
            print(f"Saved HTML to: {args.output_html}")
        except Exception as exc:
            print(f"Failed to save HTML to {args.output_html}: {exc}")

    if args.dry_run:
        print("Dry run enabled: not sending email.")
        return 0

    # Choose delivery method
    system_is_windows = platform.system().lower().startswith("win")
    use_outlook = args.use_outlook or (system_is_windows and not args.use_smtp)

    try:
        if use_outlook:
            send_email_via_outlook(subject=args.subject, html_body=html_body, to=args.to, cc=args.cc)
            print("Email sent via Outlook.")
        else:
            send_email_via_smtp(
                subject=args.subject,
                html_body=html_body,
                to=args.to,
                cc=args.cc,
                smtp_host=args.smtp_host,
                smtp_port=args.smtp_port,
                smtp_user=args.smtp_user,
                smtp_password=args.smtp_password,
                smtp_use_starttls=args.smtp_starttls,
                mail_from=args.mail_from,
            )
            print("Email sent via SMTP.")
    except Exception as exc:
        print(f"Error sending email: {exc}")
        return 4

    return 0


if __name__ == "__main__":
    sys.exit(main())

