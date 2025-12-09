#!/usr/bin/env python3
"""
Convert a .ics file that contains working‑hours events into a CSV file.

The CSV will have the following columns:
    Date, Start, End, Summary, Duration (minutes)

Usage:
    python ics_to_csv.py working_hours.ics -o working_hours.csv
"""

import argparse
import csv
from datetime import datetime, timedelta
from pathlib import Path

from icalendar import Calendar, Event


def load_calendar(ics_path: Path) -> Calendar:
    """Read an .ics file and return a Calendar object."""
    with open(ics_path, "r", encoding="utf-8") as f:
        return Calendar.from_ical(f.read())


def event_to_row(event: Event) -> tuple[str, str, str, str, int] | None:
    """Convert a VEVENT to a CSV row.
    Returns None if the event does not have a start/end datetime."""
    dtstart = event.get("DTSTART")
    dtend = event.get("DTEND")

    if not dtstart or not dtend:
        return None

    # icalendar may return date or datetime objects
    if isinstance(dtstart.dt, datetime):
        start_dt = dtstart.dt
    else:  # plain date (all‑day event)
        start_dt = datetime.combine(dtstart.dt, datetime.min.time())

    if isinstance(dtend.dt, datetime):
        end_dt = dtend.dt
    else:
        end_dt = datetime.combine(dtend.dt, datetime.min.time())

    # If the event is all‑day (no time part), skip it
    if start_dt.time() == datetime.min.time() and end_dt.time() == datetime.min.time():
        return None

    date_str = start_dt.date().isoformat()
    start_str = start_dt.time().strftime("%H:%M")
    end_str = end_dt.time().strftime("%H:%M")
    summary = event.get("SUMMARY", "")

    # Duration in minutes
    duration = int((end_dt - start_dt).total_seconds() // 60)

    return date_str, start_str, end_str, summary, duration


def parse_events(cal: Calendar) -> list[tuple[str, str, str, str, int]]:
    """Return a list of rows extracted from all VEVENTs in the calendar."""
    rows = []
    for component in cal.walk("VEVENT"):
        row = event_to_row(component)
        if row:
            rows.append(row)
    return rows


def write_csv(rows: list[tuple[str, str, str, str, int]], out_path: Path) -> None:
    """Write the rows to a CSV file."""
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Date", "Start", "End", "Summary", "Duration(min)"])
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="ICS → CSV for working hours")
    parser.add_argument(
        "ics_file",
        type=Path,
        help="Path to the .ics file containing your working‑hours events",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("working_hours.csv"),
        help="Output CSV file (default: working_hours.csv)",
    )
    args = parser.parse_args()

    cal = load_calendar(args.ics_file)
    rows = parse_events(cal)

    if not rows:
        print(f"Warning: No events found in {args.ics_file}")
    else:
        write_csv(rows, args.output)
        print(f"Converted {len(rows)} events to {args.output}")


if __name__ == "__main__":
    main()
