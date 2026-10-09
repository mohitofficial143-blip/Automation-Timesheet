# Automation Timesheet

Automated monthly Excel timesheet generation and email delivery.

## Setup

### 1. Configure Leave Data

Edit `config.yaml`:
```yaml
employee:
  name: "Your Name"
  id: "YOUR_ID"
  client: "Company"
  location: "City"

timesheets:
  - name: "Timesheet_1"
    start: "2026-08-01"
    end: "2026-08-31"
  - name: "Timesheet_2"
    start: "2026-07-26"
    end: "2026-08-26"

leaves:
  - date: "2026-08-06"
    type: "Earned Leave"
```

### 2. Run Generator

```bash
python timesheet_generator.py --dry-run
```

Output: `Timesheet_AUG_2026.xlsx`, `Timesheet_JUL_2026.xlsx`

### 3. Send Emails (Production)

Set these environment variables:
```bash
export SMTP_HOST=smtp.gmail.com
export SMTP_USER=your.email@gmail.com
export SMTP_PASS=your-app-password
```

Then run:
```bash
python timesheet_generator.py --email
```

## Gmail App Password Setup

1. Go to https://myaccount.google.com/apppasswords
2. Select "Mail" and "Other (Custom name)"
3. Enter name: "Automation-Timesheet"
4. Click "Generate"
5. Copy the 16-character password (no spaces)
6. Use this as `SMTP_PASS`

**Note:** You need 2-Step Verification enabled on your Google account first.

## GitHub Actions

The workflow runs automatically on the 1st of each month at 9:00 AM IST.

To add secrets:
1. Go to repo → Settings → Secrets and variables → Actions
2. Add:
   - `SMTP_HOST`: smtp.gmail.com
   - `SMTP_USER`: your.email@gmail.com
   - `SMTP_PASS`: your-app-password

## Test

```bash
pytest tests/ -v
```