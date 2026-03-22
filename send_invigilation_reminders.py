#!/usr/bin/env python3
"""Send invigilation reminder emails from the UG worksheet.

This script reads the UG sheet from the workbook, matches invigilator names to email
addresses in a CSV file, and sends reminders for exams occurring a configurable
number of days in the future.
"""

from __future__ import annotations

import argparse
import csv
import os
import smtplib
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile

WORKBOOK_DEFAULT = "2025_2026 Invigilation (Semester 1).xlsx"
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


@dataclass(frozen=True)
class Assignment:
    exam_date: date
    day_label: str
    time_label: str
    course: str
    venue: str
    invigilators: tuple[str, ...]


def col_to_num(col: str) -> int:
    value = 0
    for char in col:
        if char.isalpha():
            value = value * 26 + ord(char.upper()) - 64
    return value


def load_shared_strings(zip_file: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("main:si", NS):
        strings.append("".join(text.text or "" for text in item.iterfind(".//main:t", NS)))
    return strings


def locate_sheet(zip_file: ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    relationships = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall("pkgrel:Relationship", NS)
    }
    for sheet in workbook.find("main:sheets", NS):
        if sheet.attrib["name"].lower() == sheet_name.lower():
            rel_id = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            return f"xl/{rel_map[rel_id]}"
    raise ValueError(f"Worksheet '{sheet_name}' was not found in the workbook.")


def load_sheet_cells(path: Path, sheet_name: str) -> dict[int, dict[int, str | None]]:
    with ZipFile(path) as zip_file:
        shared = load_shared_strings(zip_file)
        sheet_path = locate_sheet(zip_file, sheet_name)
        root = ET.fromstring(zip_file.read(sheet_path))

    data: dict[int, dict[int, str | None]] = defaultdict(dict)
    for cell in root.findall(".//main:sheetData/main:row/main:c", NS):
        ref = cell.attrib.get("r", "")
        col = "".join(ch for ch in ref if ch.isalpha())
        row = "".join(ch for ch in ref if ch.isdigit())
        if not row:
            continue
        row_index = int(row)
        col_index = col_to_num(col)
        value_node = cell.find("main:v", NS)
        cell_type = cell.attrib.get("t")
        if cell_type == "s" and value_node is not None:
            value = shared[int(value_node.text)]
        elif cell_type == "inlineStr":
            value = "".join(text.text or "" for text in cell.iterfind(".//main:t", NS))
        elif value_node is not None:
            value = value_node.text
        else:
            value = None
        data[row_index][col_index] = value
    return data


def parse_exam_date(day_label: str) -> date:
    parts = day_label.split()
    if len(parts) < 2:
        raise ValueError(f"Could not parse exam date from '{day_label}'.")
    return datetime.strptime(parts[-1], "%d/%m/%Y").date()


def parse_assignments(path: Path, sheet_name: str = "UG") -> list[Assignment]:
    data = load_sheet_cells(path, sheet_name)
    assignments: list[Assignment] = []
    current_day = current_time = current_course = None

    for row_index in sorted(data):
        if row_index < 14:
            continue

        row = data[row_index]
        day_label = row.get(2)
        time_label = row.get(3)
        course = row.get(4)
        venue = row.get(5)
        invigilators = row.get(6)

        if not any(value is not None for value in (day_label, time_label, course, venue, invigilators)):
            continue

        if day_label:
            current_day = day_label.strip()
        if time_label:
            current_time = time_label.strip()
        if course:
            current_course = course.strip()

        if venue and invigilators:
            invigilator_list = tuple(
                item.strip() for item in invigilators.split(",") if item.strip()
            )
            assignments.append(
                Assignment(
                    exam_date=parse_exam_date(current_day),
                    day_label=current_day,
                    time_label=current_time,
                    course=current_course,
                    venue=venue.strip(),
                    invigilators=invigilator_list,
                )
            )

    return assignments


def load_invigilator_emails(path: Path) -> dict[str, str]:
    mapping: dict[str, str] = {}
    with path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        required = {"invigilator_name", "email"}
        if not required.issubset(reader.fieldnames or set()):
            raise ValueError(
                f"CSV file must contain headers: {', '.join(sorted(required))}."
            )
        for row in reader:
            name = (row.get("invigilator_name") or "").strip()
            email = (row.get("email") or "").strip()
            if name and email:
                mapping[name] = email
    return mapping


def build_email(sender: str, recipient: str, days_before: int, assignments: Iterable[Assignment]) -> EmailMessage:
    grouped = sorted(assignments, key=lambda item: (item.exam_date, item.time_label, item.course, item.venue))
    first_exam = grouped[0]
    subject = f"Reminder: Invigilation duty for {first_exam.course} on {first_exam.exam_date:%d %b %Y}"

    lines = [
        f"Dear {recipient.split('@')[0]},",
        "",
        f"This is your automated reminder that you have invigilation duties in {days_before} day(s).",
        "",
        "Your assigned sessions:",
    ]
    for assignment in grouped:
        lines.append(
            f"- {assignment.day_label} | {assignment.time_label} | {assignment.course} | {assignment.venue}"
        )
    lines.extend(
        [
            "",
            "Please make the necessary preparations and arrive on time.",
            "",
            "Best regards,",
            "Invigilation Reminder System",
        ]
    )

    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content("\n".join(lines))
    return message


def send_messages(sender: str, app_password: str, messages: Iterable[EmailMessage]) -> None:
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        for message in messages:
            server.send_message(message)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", default=WORKBOOK_DEFAULT, help="Path to the Excel workbook.")
    parser.add_argument(
        "--sheet",
        default="UG",
        help="Worksheet name that contains invigilation assignments (default: UG).",
    )
    parser.add_argument(
        "--contacts",
        default="invigilators_template.csv",
        help="CSV file mapping invigilator names to email addresses.",
    )
    parser.add_argument(
        "--days-before",
        type=int,
        default=1,
        help="Send reminders for exams this many days in the future.",
    )
    parser.add_argument(
        "--target-date",
        help="Override today's date (YYYY-MM-DD) for testing the reminder selection.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print emails that would be sent without connecting to Gmail.",
    )
    args = parser.parse_args()

    workbook = Path(args.workbook)
    contacts_path = Path(args.contacts)
    sender = os.environ.get("GMAIL_SENDER")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")

    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")
    if not contacts_path.exists():
        raise SystemExit(f"Contacts CSV not found: {contacts_path}")

    assignments = parse_assignments(workbook, args.sheet)
    contact_map = load_invigilator_emails(contacts_path)

    base_date = (
        datetime.strptime(args.target_date, "%Y-%m-%d").date()
        if args.target_date
        else date.today()
    )
    reminder_date = base_date + timedelta(days=args.days_before)

    by_recipient: dict[str, list[Assignment]] = defaultdict(list)
    missing_contacts: set[str] = set()

    for assignment in assignments:
        if assignment.exam_date != reminder_date:
            continue
        for invigilator in assignment.invigilators:
            email = contact_map.get(invigilator)
            if email:
                by_recipient[email].append(assignment)
            else:
                missing_contacts.add(invigilator)

    if missing_contacts:
        print("No email address configured for:")
        for name in sorted(missing_contacts):
            print(f"- {name}")
        print()

    if not by_recipient:
        print(f"No reminders to send for exams on {reminder_date:%Y-%m-%d}.")
        return 0

    messages = [
        build_email(sender or "your_gmail@gmail.com", email, args.days_before, items)
        for email, items in sorted(by_recipient.items())
    ]

    if args.dry_run:
        print(f"Dry run: {len(messages)} reminder(s) would be sent for exams on {reminder_date:%Y-%m-%d}.")
        for message in messages:
            print("=" * 72)
            print(f"To: {message['To']}")
            print(f"Subject: {message['Subject']}")
            print(message.get_content())
        return 0

    if not sender or not app_password:
        raise SystemExit(
            "Set GMAIL_SENDER and GMAIL_APP_PASSWORD before sending live emails."
        )

    send_messages(sender, app_password, messages)
    print(f"Sent {len(messages)} reminder(s) for exams on {reminder_date:%Y-%m-%d}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
