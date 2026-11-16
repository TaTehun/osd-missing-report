# OS&D Missing Report

A weekly Python automation pipeline that downloads OS&D (Over, Short & Damage) reports and IBD Origin files from email, compares them to identify missing items, and distributes plant-level missing reports via email, with no manual steps.

> **Visual overview:** [project-overview.html](project-overview.html)

---

## Overview

Each week, plant teams submit OS&D reports to a shared mailbox, and an IBD Origin file is prepared by the inbound team. This pipeline:

1. **Downloads** the previous and current week's OS&D report attachments and IBD Origin file from the POP3 mailbox
2. **Validates** each downloaded OS&D file for required columns and data integrity
3. **Compares** IBD Origin vs OS&D reports per plant: items in IBD but not in OS&D are flagged as **Missing** (QC Grade A / LS excluded)
4. **Checks** Missing items against the OVG (Over/Good) file from two weeks ago, deducting items found in OVG
5. **Resolves** the master missing list: items whose serial shows up in this week's OS&D are removed
6. **Saves** final Missing items as plant-level Excel reports and adds them to the weekly master file
7. **Writes** this week's OVG file: unmatched OS&D items are carried forward; OVG and master records older than 15 weeks are pruned
8. **Distributes** plant-level missing reports and a status summary via SMTP email

---

## Project Structure

```
osd-missing-report/
├── main.py                        # Core processing: IBD vs OSD comparison, OVG, report generation
├── report_downloader.py           # POP3 downloader entry point + file validation
├── run_report_downloader.bat      # Task Scheduler trigger for downloader
├── requirements.txt
├── email_pipeline/
│   ├── email_setup.py             # All settings (paths, SMTP, plant recipients)
│   ├── downloader.py              # POP3 attachment downloader (OS&D + IBD)
│   └── sender.py                  # SMTP sender (status report, plant report, failure alert)
├── utils/
│   ├── date_utils.py              # Week/date formatting utilities
│   ├── log_utils.py               # Logger setup + log snippet extraction
│   └── pivot_utils.py             # Pivot table refresh + HTML rendering
├── study/                         # Annotated study files for each module
└── brd/                           # Business requirement documents
```

---

## Data Sources

| Source | Description |
|--------|-------------|
| OS&D Reports | One Excel file per plant per week; Summary sheet with O/S/D column |
| IBD Origin | One Excel/xlsb file per week; IB Damage data with Material, Serial number, Quantity |

> If an OS&D file is missing for a plant, that plant is skipped and reported in the status email.

---

## Output

All files are organized by week:

```
Desktop/osd weekly report/
├── data/
│   ├── osd report/                # Downloaded OS&D report files (kept 5 weeks)
│   ├── IB Damage/                 # Downloaded IBD Origin files (kept 5 weeks)
│   ├── OVG/
│   │   └── OSD_OVG_W21.xlsx       # One OVG file per week (15 weeks rolling)
│   └── master/
│       └── W21_missing_master.xlsx  # Open missing items across weeks (15 weeks rolling)
├── report/
│   └── W21/
│       ├── W21 PLANT1 Missing items from OSD Report.xlsx
│       ├── W21 PLANT2 Missing items from OSD Report.xlsx
│       └── ...
└── logs/
    └── main_20260528.log
```

---

## How It Works

### 1. Download (`report_downloader.py`)
Connects to the POP3 mailbox and scans emails newest-first, stopping at emails older than today. Matches OS&D reports by subject tokens (`OS`, `Report`, week number, plant code) for both the previous and current week, skipping subjects that contain "Missing" (the pipeline's own report threads). IBD Origin is matched by `IB Damage Origin` + week number. OS&D attachments are renamed to `{week} {plant} OS&D Report` so `main.py` finds them regardless of the sender's file name.

After downloading, each OS&D file is validated:
- A sheet named Summary must exist (any casing)
- `O/S/D`, `Model`, and `Qty` must all be in the first row; every missing header is reported together
- At least one `D` value must exist in the O/S/D column

### 2. IBD vs OSD Comparison (`main.py → find_missing`)
Friday runs process the current week; other days process the previous week. IBD rows with QC Grade `A` or `LS` are dropped first. For each plant, IBD rows are compared against OS&D D-rows using `_match_and_consume`:

| Condition | Matching Method |
|-----------|----------------|
| IBD has a valid Serial Number (14–15 alphanumeric) | 1:1 match against OSD `Serial Number(s)`; if none, retried with a leading `O` in the OSD serial read as `0` (data-entry fix, reported in the status email) |
| IBD has no Serial Number | Match by `Material` vs OSD `Model`; OSD qty is summed and deducted |

IBD rows with remaining qty after matching → **Missing candidates**
OSD D-rows not consumed by any IBD row → **OVG new** (OS&D has it, IBD doesn't)

### 3. OVG Check (`main.py → check_ovg`)
Missing candidates are checked against the OVG file from two weeks ago (`OSD_OVG_{week}.xlsx`). Items found in OVG are deducted from Missing using the same `_match_and_consume` logic. What remains after deduction is the **final Missing**.

### 4. Master Resolve (`main.py → resolve_master`)
The master file from two weeks ago holds open missing items. Any item whose serial appears in this week's OS&D D-rows is removed; this week's final Missing is appended, and the result is saved as `{week}_missing_master.xlsx`.

### 5. OVG Update (`main.py → save_ovg`)
- Remaining OVG (not matched by this week's Missing) is kept
- New OVG rows from this week's comparison are appended
- Records older than 15 weeks are pruned automatically
- OSD columns are preserved exactly: `Plant, Week, Date, Container, SN, Live/Drop, Seal in Tact, O/S/D, Model, SN QTY, Qty, Serial Number(s), WH Note`

### 6. Distribution (`sender.py`)
- **Status summary email**: sent to the configured list with per-plant Success/Fail status, O→0 corrected serials, and OVG report notice
- **Plant-level report emails**: each plant's missing report is attached and sent to plant-specific recipients
- **Failure alert**: sent if an unexpected error occurs during processing, with log snippet attached

---

## Configuration

All settings are managed in `email_pipeline/email_setup.py`:

| Setting | Description |
|---------|-------------|
| `OSD_DIR` / `IBD_DIR` / `OVG_DIR` / `MASTER_DIR` | Paths to data folders |
| `PLANT_CODES` | List of all plant codes to process |
| `OSD_SUBJECT_TOKENS` / `IBD_SUBJECT_TOKENS` | Email subject match keywords |
| `SMTP_SERVER` / `SMTP_PORT` | SMTP server settings |
| `STATUS_TO` / `STATUS_CC` | Status summary email recipients |
| `PLANT_RECIPIENTS` | Per-plant report email recipients `{code: {to, cc, bcc}}` |
| `MAIN_FAILURE_TO` | Failure alert recipients |

Credentials (POP3/SMTP username and password) are stored in a `.env` file on the network drive and loaded at runtime.

---

## Warehouse → Plant Mapping

| Warehouse Code | Plant Code |
|---------------|-----------|
| WH1 | PLANT1 |
| WH2 | PLANT2 |
| WH3 | PLANT3 |
| WH4 | PLANT4 |
| WH5 | PLANT5 |
| WH6 | PLANT6 |
| WH7 | PLANT7 |
| WH8 | PLANT8 |

---

## Requirements

```
pandas==3.0.3
openpyxl==3.1.5
pyxlsb==1.0.10
pytz==2026.2
python-dotenv==1.2.2
```

---

## Running

```bash
# Download OS&D reports and IBD Origin from email
python report_downloader.py
# or double-click run_report_downloader.bat

# Process reports and generate missing item reports
python main.py
```

Intended to be triggered weekly (after OS&D reports are received) via Windows Task Scheduler.

---

## Changelog

### v1.1 · 2026-10-21 · Master list, two-week OVG, SN correction
- Weekly master missing file: open items carried across weeks and resolved when their serial shows up in OS&D
- OVG stored per week; each run checks the OVG from two weeks ago
- QC Grade A / LS rows excluded before matching
- Leading `O` → `0` serial correction, listed in the status email
- Friday runs process the current week; downloader fetches previous and current week
- OS&D attachments renamed to a fixed pattern; "Missing" reply threads skipped
- Validation checks `O/S/D`, `Model`, `Qty` headers together; Summary sheet matched in any casing
- One plant removed from processing (8 plants)

### v1.0 · 2026-05-28 · Initial Release
- POP3 downloader for OS&D reports + IBD Origin
- IBD vs OSD comparison with Serial Number and Material+Qty matching
- OVG accumulation file (15-week rolling, OSD column schema)
- Plant-level missing report generation and SMTP distribution
- File validation on download (Summary sheet, O/S/D column, D values)
- Automatic cleanup of files older than 5 weeks
