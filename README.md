# Receipt Invoice and Email Automation

This tool reads the student contribution Excel file, generates a PDF receipt for each student using the provided receipt template, and optionally emails the receipt to each student.

## Setup

1. Create a virtual environment and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Copy the example environment file and fill in your SMTP credentials:

```bash
cp .env.example .env
```

> **Note:** Use an app password for Gmail accounts.

## Usage

Generate receipts and preview emails without sending:

```bash
python receipt_emailer.py --excel "Student Funds List 2026-2027.xlsx" --template "Receipt Template.pdf" --output receipts
```

Send emails (after you confirm the generated PDFs look correct):

```bash
python receipt_emailer.py --excel "Student Funds List 2026-2027.xlsx" --template "Receipt Template.pdf" --output receipts --send
```

## Customizing the Receipt Layout

The receipt coordinates live in `template_positions.json`. The values are defined as percentages of the PDF page. If text is misaligned, adjust the `x_pct`, `y_pct`, or `max_width` values.

## Files

- `receipt_emailer.py` – main script.
- `receipt_config.json` – sender details, semester name, and email subject.
- `email_template.txt` – editable email template.
- `template_positions.json` – placement of fields on the receipt PDF.

