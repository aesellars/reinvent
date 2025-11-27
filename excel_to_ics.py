import argparse
import re
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from dateutil import parser as date_parser
from icalendar import Alarm, Calendar, Event
from pytz import timezone


REQUIRED_COLUMNS = {"title", "date", "time", "venue"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert Excel rows into individual iCalendar (.ics) files "
            "with travel-time friendly metadata for Apple Calendar."
        )
    )
    parser.add_argument("excel", type=Path, help="Path to the Excel file containing events")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("ics_output"),
        help="Directory where .ics files will be written",
    )
    parser.add_argument(
        "--timezone",
        default="UTC",
        help="IANA timezone name to localize event times (default: UTC)",
    )
    parser.add_argument(
        "--alert-minutes",
        type=int,
        default=30,
        help="Minutes before start to trigger travel alert (default: 30)",
    )
    return parser.parse_args()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {col: col.strip().lower() for col in df.columns}
    df = df.rename(columns=mapping)
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")
    return df


def parse_datetime(value: Any, tz) -> datetime:
    if isinstance(value, datetime):
        dt = value
    else:
        dt = date_parser.parse(str(value))
    if dt.tzinfo is None:
        dt = tz.localize(dt)
    else:
        dt = dt.astimezone(tz)
    return dt


def parse_time_range(value: Any) -> tuple[time, time]:
    parts = re.split(r"\s*-\s*", str(value), maxsplit=1)
    if len(parts) != 2:
        raise ValueError(
            "Time column must contain a start and end time separated by a hyphen (e.g., '9:00 AM - 10:00 AM')."
        )
    start_time = date_parser.parse(parts[0]).time()
    end_time = date_parser.parse(parts[1]).time()
    return start_time, end_time


def row_description(row: Dict[str, Any]) -> str:
    lines = ["Event details from spreadsheet:"]
    for key, value in row.items():
        lines.append(f"- {key}: {value}")
    return "\n".join(lines)


def safe_filename(title: str, row_index: int) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", title).strip("_") or "event"
    return f"{safe}_{row_index}.ics"


def build_event(row: Dict[str, Any], tz, alert_minutes: int) -> Calendar:
    title = str(row.get("title", "Event")).strip() or "Event"
    venue = str(row.get("venue", "")).strip()

    start = parse_datetime(row["date"], tz)
    start_time, end_time = parse_time_range(row["time"])

    dtstart = start.replace(hour=start_time.hour, minute=start_time.minute, second=start_time.second, microsecond=0)
    dtend = start.replace(hour=end_time.hour, minute=end_time.minute, second=end_time.second, microsecond=0)
    if dtend <= dtstart:
        dtend = dtstart + timedelta(hours=1)

    cal = Calendar()
    cal.add("prodid", "-//Excel to ICS Converter//EN")
    cal.add("version", "2.0")

    event = Event()
    event.add("summary", title)
    event.add("dtstart", dtstart)
    event.add("dtend", dtend)
    if venue:
        event.add("location", venue)
    event.add("description", row_description(row))
    event.add("transp", "OPAQUE")

    # Enable Apple Calendar travel-time handling.
    event.add("x-apple-travel-advisory-behavior", "AUTOMATIC")
    event.add("x-apple-travel-duration", "PT0M")

    alert = Alarm()
    alert.add("action", "DISPLAY")
    alert.add("description", "Leave now to arrive on time.")
    alert.add("trigger", timedelta(minutes=-abs(alert_minutes)))
    alert.add("x-apple-local-default-alarm", True)
    alert.add("x-apple-default-alarm", True)
    event.add_component(alert)

    cal.add_component(event)
    return cal


def main():
    args = parse_args()
    tz = timezone(args.timezone)

    df = pd.read_excel(args.excel)
    df = normalize_columns(df)

    args.output.mkdir(parents=True, exist_ok=True)

    for idx, row in df.iterrows():
        cal = build_event(row.to_dict(), tz, args.alert_minutes)
        filename = safe_filename(str(row.get("title", "event")), idx + 1)
        output_path = args.output / filename
        with open(output_path, "wb") as f:
            f.write(cal.to_ical())
        print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
