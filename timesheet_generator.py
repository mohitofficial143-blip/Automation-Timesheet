"""
Automation Timesheet Generator

Reads config.yaml, generates two monthly Excel timesheets, and emails them.
Supports: employee details, date ranges, leave types (Earned Leave, Sick Leave, etc.),
office holidays, weekend detection, Total Leaves summary row with count + reasons.

Usage:
    python timesheet_generator.py [--config config.yaml] [--dry-run]
"""

import calendar
import datetime
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import yaml


# ---------------------------------------------------------------------------
# Constants (matches reference Excel exactly)
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFFFF", bold=True, name="Calibri", size=11)
WEEKEND_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
HOLIDAY_FILL = PatternFill(start_color="FFFCE4D6", end_color="FFFCE4D6", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin", color="FF000000"),
    right=Side(style="thin", color="FF000000"),
    top=Side(style="thin", color="FF000000"),
    bottom=Side(style="thin", color="FF000000"),
)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
COLUMN_WIDTHS = {"A": 15, "B": 15, "C": 18, "D": 18, "E": 12, "F": 20}
HEADERS = ["WEEKDAY", "DATE", "LOG-IN TIMINGS", "LOG-OUT TIMINGS", "Total Hours", "Remarks"]
EMPLOYEE_FIELDS = ["EMPLOYEE NAME", "EMPLOYEE ID", "CLIENT", "LOCATION"]
LOG_IN = "9:30:00 AM"
LOG_OUT_FRI = "5:30:00 PM"
LOG_OUT_WORKDAY = "6:30:00 PM"
HOURS_FRI = 8
HOURS_WORKDAY = 9
WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def load_config(config_path: str = "config.yaml") -> dict:
    """Load and validate config.yaml."""
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    with open(path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # Basic validation
    required = {"employee", "timesheets", "leaves"}
    missing = required - set(cfg.keys())
    if missing:
        raise ValueError(f"Missing config keys: {missing}")
    for ts in cfg["timesheets"]:
        for key in ("name", "start", "end"):
            if key not in ts:
                raise ValueError(f"Timesheet missing '{key}': {ts}")
    return cfg


def build_timesheet(cfg: dict, ts_cfg: dict) -> tuple[openpyxl.Workbook, str]:
    """
    Build one timesheet workbook from config.

    Returns:
        (wb, output_path) where output_path is the full file path.
    """
    start = datetime.date.fromisoformat(ts_cfg["start"])
    end = datetime.date.fromisoformat(ts_cfg["end"])
    emp = cfg["employee"]
    leaves_by_date = {datetime.date.fromisoformat(l["date"]): l for l in cfg["leaves"]}

    wb = openpyxl.Workbook()
    ws = wb.active
    month_label = calendar.month_name[start.month][:3].upper()
    ws.title = month_label

    # --- Styles ---
    # Employee info rows 2-5
    for r in range(2, 6):
        ws.cell(row=r, column=1).font = Font(bold=True, name="Calibri", size=11)

    # --- Row 1: Title (merged) ---
    ws.merge_cells("A1:F1")
    ws["A1"] = f"TIMESHEET FOR THE PERIOD '{month_label}-{start.year}'"
    ws["A1"].font = Font(size=14, bold=True, name="Calibri")
    ws["A1"].alignment = CENTER_ALIGN

    # --- Rows 2-5: Employee info ---
    info_map = {
        "EMPLOYEE NAME": emp.get("name", ""),
        "EMPLOYEE ID": emp.get("id", ""),
        "CLIENT": emp.get("client", ""),
        "LOCATION": emp.get("location", ""),
    }
    for i, label in enumerate(EMPLOYEE_FIELDS, start=2):
        ws.cell(row=i, column=1).value = label
        ws.cell(row=i, column=2).value = info_map[label]

    # --- Row 7: Headers ---
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=7, column=col)
        cell.value = header
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # --- Data rows (8+) ---
    row = 8
    leave_reasons: list[str] = []
    current_date = start
    while current_date <= end:
        weekday_num = current_date.weekday()  # 0=Mon, 6=Sun
        weekday_name = WEEKDAY_NAMES[weekday_num]
        date_str = current_date.strftime("%d-%m-%Y")

        fill = None
        log_in = ""
        log_out = ""
        hours = ""
        remark = ""

        if current_date in leaves_by_date:
            leave_info = leaves_by_date[current_date]
            fill = HOLIDAY_FILL
            remark = leave_info["type"].lower()  # e.g. "Medical leave"
            leave_reasons.append(leave_info["type"])
        elif weekday_num >= 5:  # Saturday or Sunday
            fill = WEEKEND_FILL
            remark = "Weekoff"
        else:
            log_in = LOG_IN
            if weekday_num == 4:  # Friday
                log_out = LOG_OUT_FRI
                hours = HOURS_FRI
            else:
                log_out = LOG_OUT_WORKDAY
                hours = HOURS_WORKDAY

        data = [weekday_name, date_str, log_in, log_out, hours, remark]
        for col, value in enumerate(data, start=1):
            cell = ws.cell(row=row, column=col)
            cell.value = value
            cell.border = THIN_BORDER
            cell.alignment = CENTER_ALIGN
            if fill:
                cell.fill = fill

        row += 1
        current_date += datetime.timedelta(days=1)

    # --- Total Leaves summary row ---
    total_leaves_row = row
    ws.cell(row=total_leaves_row, column=1).value = "Total Leaves"
    ws.cell(row=total_leaves_row, column=1).font = Font(bold=True, name="Calibri", size=11)
    ws.cell(row=total_leaves_row, column=2).value = len(leave_reasons)
    if leave_reasons:
        ws.cell(row=total_leaves_row, column=3).value = ", ".join(leave_reasons)

    # --- Column widths ---
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # --- Save ---
    output_dir = Path(__file__).parent
    filename = f"Timesheet_{month_label}_{start.year}.xlsx"
    save_path = output_dir / filename
    wb.save(save_path)
    return wb, str(save_path)


def send_email(from_addr: str, to_addr: str, subject: str, body: str, attachments: list[str]):
    """Send email with optional attachments via SMTP."""
    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    for path in attachments:
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
        msg.attach(part)
    # Send via SMTP (defaults to localhost; adjust in production)
    try:
        with smtplib.SMTP("localhost") as s:
            s.send_message(msg)
    except ConnectionRefusedError:
        # Fallback: print what would be sent (for dry-run / dev)
        print(f"[SMTP unavailable — mail would be sent to {to_addr}]")
        print(f"  Subject: {subject}")
        for a in attachments:
            print(f"  Attachment: {a}")


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Generate monthly timesheets and email them.")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--dry-run", action="store_true", help="Generate files without emailing")
    args = parser.parse_args()

    try:
        cfg = load_config(args.config)
    except (FileNotFoundError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    outputs: list[str] = []
    for ts in cfg["timesheets"]:
        try:
            _, path = build_timesheet(cfg, ts)
            outputs.append(path)
            print(f"Generated: {path}")
        except Exception as e:
            print(f"ERROR generating {ts['name']}: {e}", file=sys.stderr)

    if not outputs:
        print("ERROR: No timesheets generated.", file=sys.stderr)
        sys.exit(1)

    if args.dry_run:
        print(f"\n[Dry-run mode — skipped email. {len(outputs)} file(s) generated.]")
        return

    # Email
    recipient = "mohit.official143@gmail.com"
    from_addr = recipient  # same as recipient for dev (set to SMTP auth in prod)
    # Subject/body per timesheet month
    for path in outputs:
        month_label = Path(path).stem.replace("Timesheet_", "").split("_")[0]
        subject = f"Timesheets for {month_label} — {cfg['employee']['name']}"
        body = f"timesheet of {month_label} — {cfg['employee']['name']}"
        send_email(from_addr, recipient, subject, body, [path])
        print(f"Emained: {path} -> {recipient}")


if __name__ == "__main__":
    main()