# Automation Timesheet — Project Specification

**Issue:** MOHISPACE-4 | **Status:** Phase 0 Approved  
**Author:** Mohit Banra | **QA Reviewer:** QA Testing Agent  
**Last updated:** 2026-10-09

---

## 1. Objective

Automate the manual process of generating two monthly Excel timesheets and emailing them — one scheduled execution, two outputs, zero manual steps.

---

## 2. Scope

| Item | Detail |
|------|--------|
| Outputs | 2 Excel `.xlsx` files, one per timesheet range |
| Recipient | `mohit.official143@gmail.com` |
| Scheduling | GitHub Actions cron (to be confirmed) |
| Manual input required | None (all driven by `config.yaml`) |

---

## 3. Timesheet Ranges

| Timesheet | Start Date | End Date | Output Filename |
|-----------|-----------|----------|----------------|
| Timesheet 1 | 01 Aug 2026 | 31 Aug 2026 | `Timesheet_Aug_2026.xlsx` |
| Timesheet 2 | 26 Jul 2026 | 26 Aug 2026 | `Timesheet_Jul_2026.xlsx` |

> File naming convention: `Timesheet_{Month-short}_{YYYY}.xlsx`  
> Month short = first 3 letters, uppercase (Jan, Feb, …, Dec)

---

## 4. Excel Structure (per sheet)

### 4.1 Sheet Metadata
- **Sheet name:** First 3 letters of the start month, uppercase (e.g. `JUL`, `AUG`)

### 4.2 Rows 1–5
| Row | A | B | Notes |
|-----|---|---|-------|
| 1 (merged A1:F1) | `TIMESHEET FOR THE PERIOD 'Sept-2026'` | — | Month-label format, not date-range |
| 2 | `EMPLOYEE NAME` | `{employee.name}` | |
| 3 | `EMPLOYEE ID` | `{employee.id}` | |
| 4 | `CLIENT` | `{employee.client}` | |
| 5 | `LOCATION` | `{employee.location}` | |

### 4.3 Row 7 — Column Headers
```
A: WEEKDAY   B: DATE       C: LOG-IN TIMINGS   D: LOG-OUT TIMINGS   E: Total Hours   F: Remarks
```
- Background: `#1F4E78` (dark blue, opaque `FF1F4E78`)
- Font: white (`#FFFFFF`), bold, Calibri 11
- Alignment: center / center
- Border: thin black on all sides

### 4.4 Data Rows (starting row 8)

| Day Type | Fill Color | Log-In | Log-Out | Hours | Remarks |
|----------|-----------|--------|---------|-------|---------|
| Weekend (Sat/Sun) | `#D9D9D9` | — | — | — | `Weekoff` |
| Leave (from config) | `#FCE4D6` | — | — | — | Leave type (e.g. `Medical leave`, `Earned Leave`) |
| Office Holiday | `#FCE4D6` | — | — | — | `Office Holiday` |
| Mon–Thu | None (white) | `9:30:00 AM` | `6:30:00 PM` | `9` | — |
| Fri | None (white) | `9:30:00 AM` | `5:30:00 PM` | `8` | — |

> Leave rows take priority over weekend/holiday coloring when a date matches a configured leave.

### 4.5 Bottom Summary Row
```
A: Total Leaves    B: <count of leave-type rows>    C-F: empty
```
Example: `Total Leaves | 1`

No additional detail row beyond this — leave reasons are already in the `Remarks` column per row.

### 4.6 Column Widths
```
A: 15  B: 15  C: 18  D: 18  E: 12  F: 20
```

---

## 5. Input Configuration (`config.yaml`)

```yaml
employee:
  name: "Mohit Banra"
  id: "FFE520"
  client: "Partnr"
  location: "Gurugram"

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
  - date: "2026-08-27"
    type: "Medical Leave"
```

### Fields
| Field | Type | Required | Notes |
|-------|------|----------|-------|
| `employee.name` | string | yes | Shown in row 2 |
| `employee.id` | string | yes | Shown in row 3 |
| `employee.client` | string | yes | Shown in row 4 |
| `employee.location` | string | yes | Shown in row 5 |
| `timesheets[*].name` | string | yes | Used for logging; no functional impact |
| `timesheets[*].start` | date (ISO) | yes | Inclusive |
| `timesheets[*].end` | date (ISO) | yes | Inclusive |
| `leaves[*].date` | date (ISO) | yes | Must fall within at least one timesheet range |
| `leaves[*].type` | string | yes | Displayed in Remarks column (lowercased in cell e.g. `Medical leave`) |

---

## 6. Email Specification

| Field | Template |
|-------|----------|
| To | `mohit.official143@gmail.com` |
| CC | — (not required) |
| BCC | — (not required) |
| Subject | `Timesheets for {month} — {employee.name}` |
| Body | `timesheet of {month} — {employee.name}` |
| Attachments | Both `.xlsx` files in one email |

> `{month}` = month label from the timesheet (e.g. `Aug`, `Jul`)

---

## 7. Error Handling & Notification

- If configuration file is missing or malformed → **send error email** with subject `ERROR: Timesheet generation failed — {reason}`, body explains the issue
- If Excel generation fails for any timesheet → **skip that timesheet**, continue generating others, attach only successful files
- If email send fails → log error, no retry (manual re-send possible)

---

## 8. GitHub Actions Scheduling

Placeholder for cron schedule (to be provided by Mohit):
```yaml
on:
  schedule:
    - cron: '0 9 1 * *'  # TBD — Mohit will confirm
  workflow_dispatch:      # manual trigger always available
```

---

## 9. Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl>=3.1.0` | Excel generation |
| `PyYAML>=6.0` | Config parsing |
| `pytest>=7.0` | Tests |

---

## 10. Validation Rules (Phase 4 QA)

| Rule | Expected |
|------|----------|
| Sheet name | First 3 chars of start month, uppercase |
| A1 title format | `TIMESHEET FOR THE PERIOD '{Month-YYYY}'` |
| Header row (7) | Exact text + dark blue fill + white bold font |
| Weekend fill | `#D9D9D9` on all 6 cells |
| Leave fill | `#FCE4D6` on all 6 cells, Remarks = leave type (title-cased) |
| Friday hours | 8, log-out = `5:30:00 PM` |
| Mon–Thu hours | 9, log-out = `6:30:00 PM` |
| Total Leaves row | Present at last non-empty row, count matches leave-type days |
| Column widths | A=15, B=15, C=18, D=18, E=12, F=20 |
| Merge A1:F1 | Present |
| No extra rows after summary | Empty rows not written |

---

## 11. Files & Directory Structure (Target)

```
Automation-Timesheet/
├── SPEC.md              # This document
├── config.yaml          # Runtime config (template below)
├── requirements.txt     # Python dependencies
├── timesheet_generator.py  # Main script
├── tests/
│   ├── __init__.py
│   ├── test_generator.py
│   └── test_config.py
└── .github/workflows/
    └── timesheet.yml    # CI workflow
```

---

*Phase 0 approved pending MOHIT sign-off.*