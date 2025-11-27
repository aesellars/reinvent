# Excel to iOS-ready ICS Converter

This command-line tool generates an individual iCalendar (`.ics`) file for each event in an Excel workbook. It is designed to support Apple Calendar travel-time alerts and to embed the full row data in the meeting notes for context.

## Features
- Reads events from an Excel sheet (any worksheet name) using column headers.
- Produces one `.ics` file per event with the title as the summary and the venue as the location.
- Includes the full row contents in the description/meeting notes.
- Adds Apple travel-time advisory metadata and a default travel alert.

## Expected columns
The script expects the following columns (case-insensitive):

- `title` – Event title.
- `date` – Event date.

- `time` – Start and end time range for the event, formatted with a hyphen (e.g., `9:00 AM - 10:00 AM`).

- `venue` – Location text for the event.

Date and time fields can be Excel date/time cells or strings; they are combined into timezone-aware datetimes using the `--timezone` option (default: `UTC`).

## Usage
1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Run the converter:

   ```bash
   python excel_to_ics.py path/to/events.xlsx --output ./ics_output --timezone "America/New_York" --alert-minutes 20
   ```

   - `--output`: Directory where the `.ics` files will be written (created if missing).
   - `--timezone`: IANA timezone name used to localize the start/end times (default `UTC`).
   - `--alert-minutes`: Minutes before start to trigger the travel-time alert (default `30`).

Each `.ics` file uses the event title for the filename (with unsafe characters stripped) and appends the row number to avoid collisions.

## Notes on travel time
Apple Calendar interprets `X-APPLE-TRAVEL-ADVISORY-BEHAVIOR` and `X-APPLE-TRAVEL-DURATION` metadata for automatic travel-time handling. The generated events also include a default alert (`VALARM`) that fires the specified number of minutes before the start time to prompt leaving for travel.
