# Invigilation-Reminder-Email

This repository contains a lightweight Python script for sending automated Gmail reminders to invigilators before their exam day.

## Files
- `send_invigilation_reminders.py` parses the `UG` worksheet from `2025_2026 Invigilation (Semester 1).xlsx`, matches invigilators to recipient emails, and sends reminders.
- `invigilators_template.csv` is the template you can fill in with the real email addresses for each invigilator.

## Setup
1. Fill in `invigilators_template.csv` with the actual email addresses.
2. Create a Gmail App Password for the Gmail account that will send reminders.
3. Export these environment variables:

```bash
export GMAIL_SENDER="your_gmail@gmail.com"
export GMAIL_APP_PASSWORD="your_16_character_app_password"
```

## Dry run
Use a dry run first to verify which reminders will be generated without sending email:

```bash
python send_invigilation_reminders.py --dry-run --target-date 2026-03-04
```

The example above previews reminders for exams scheduled on `2026-03-05`, because the script defaults to `--days-before 1`.

## Send live reminders
```bash
python send_invigilation_reminders.py
```

## Optional cron automation
To run every morning at 08:00 UTC:

```cron
0 8 * * * cd /workspace/Invigilation-Reminder-Email && /usr/bin/env python send_invigilation_reminders.py >> reminder.log 2>&1
```

## Notes
- The script uses only the Python standard library; no extra packages are required.
- If an invigilator in the workbook has no matching email in the CSV file, the script reports the missing contact and skips sending to that person.
- The current workbook contains some shortened names such as `Ephraim`, `Hajra`, `Lucas`, and `Fransica`; replace those with the correct full names in both the workbook and CSV if needed.
