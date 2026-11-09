#!/usr/bin/python
"""
Belar AMMA-2 Integrated Logger and Flask Remote Web Dashboard.

This script runs a background thread to continuously poll data from the
Belar AMMA-2 via RS-232, logs entries to data.log, appends them to a clean
Excel file (modulation_data.xlsx), and serves a live web dashboard.
"""

import logging
import os
import csv
import io
import math
import re
from datetime import datetime, timedelta
from bisect import bisect_left
from threading import Thread, Lock
import time
from flask import Flask, jsonify, render_template_string, request
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import serial # allow the communication with devices through serial port

# Initialize logging configuration immediately at the global level
logging.basicConfig(
    filename="./data.log",
    level=logging.INFO,
    format="%(asctime)s : %(message)s",
    datefmt="%Y-%m-%d %H:%M",
)

app = Flask(__name__)

# Global cache to share live telemetry between the serial background thread and Flask
live_telemetry = {
    "DA": "No Data",  # Negative Peak
    "DB": "No Data",  # Positive Peak
    "DG": "No Data",  # Carrier Reference
    "DH": "No Data",  # AM Noise / Residual
    "DN": "No Data",  # Average / RMS Modulation
    "last_updated": "Never"
}

EXCEL_FILE = "./modulation_data.xlsx"
GRMS_FILE = "./grms_comparison.csv"
AMMA_UPLOAD_FILE = "./amma2_comparison.csv"
COMPARISON_START_TIME = "19:58:00"
COMPARISON_END_TIME = "22:50:00"
excel_lock = Lock()


def get_serial_device(device_file, baudrate=9600):
    """Initialize and return a configured serial connection."""
    return serial.Serial(
        port=device_file, # serial port to connect to (eg "com3" on windows or "/dev/ttyUSB0" on Linux
        baudrate=baudrate, # communication speed bits per second
        parity=serial.PARITY_NONE, # no parity checking
        stopbits=serial.STOPBITS_ONE,
        bytesize=serial.EIGHTBITS, # send and receive 8-bits bytes
        timeout=1
    )


def init_excel_workbook():
    """Create or verify the Excel file structure with clean styling."""
    if not os.path.exists(EXCEL_FILE):
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Telemetry Logs"

            # Define clean, professional styling
            header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")  # Dark Blue
            center_align = Alignment(horizontal="center", vertical="center")

            headers = [
                "Timestamp",
                "Negative Peak (DA)",
                "Positive Peak (DB)",
                "Carrier Ref (DG)",
                "AM Noise (DH)",
                "RMS/Avg Mod (DN)"
            ]

            ws.append(headers)

            # Apply style configurations across the header row
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_align

            # Freeze headers so scrolling down logs stays clear
            ws.freeze_panes = "A2"

            wb.save(EXCEL_FILE)
            print(f"Initialized clean Excel storage workbook at: {EXCEL_FILE}")
        except Exception as e:
            logging.error("Failed to initialize Excel workbook: %s", e)


def append_to_excel(timestamp, data_dict):
    """Append a row of captured telemetry variables safely into Excel."""
    try:
        with excel_lock:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb.active

            row_data = [
                timestamp,
                data_dict.get("DA", "No Data"),
                data_dict.get("DB", "No Data"),
                data_dict.get("DG", "No Data"),
                data_dict.get("DH", "No Data"),
                data_dict.get("DN", "No Data")
            ]

            ws.append(row_data)

            # Format the newly appended row for visual alignment consistency
            last_row = ws.max_row
            data_font = Font(name="Segoe UI", size=10)
            center_align = Alignment(horizontal="center", vertical="center")

            for cell in ws[last_row]:
                cell.font = data_font
                cell.alignment = center_align

            # Dynamically auto-fit column widths to prevent ### formatting visual bugs
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

            wb.save(EXCEL_FILE)
    except Exception as e:
        logging.error("Failed writing dataset row entry to Excel file: %s", e)

def data_collection_worker():
    """Background thread logic tracking the exact getdata.py looping routine with hotfixes."""
    global live_telemetry
    device = "/dev/ttyUSB0"

    # Initialize structural layout components for Excel spreadsheet tracking
    init_excel_workbook()

    commands = ["DA\r", "DB\r", "DG\r", "DH\r", "DN\r"]
    print("Starting background serial data collection thread...")

    while True:
        try:
            print("Attempting connection to serial device...")
            with get_serial_device(device) as ser:
                print("Connection established. Beginning collection loops...")
                # Removed redundant nested 'while True' to handle interface reconnect loops correctly
                while True:
                    print("Start collecting data ...")
                    cycle_data = {}

                    for command in commands:
                        clean_cmd = command.replace("\r", "")
                        try:
                            ser.write(command.encode('ascii'))
                            time.sleep(0.1)

                            raw_data = ser.readline()
                            data = raw_data.decode('ascii', errors='ignore').strip()

                            # Log data to disk text log precisely like the standalone script
                            logging.info(";Data for %s :; %s ", clean_cmd, data)

                            # Save to local block for Excel execution step
                            cycle_data[clean_cmd] = data if data else "No Data"

                            # Thread-safe update to the global web server cache
                            if data:
                                live_telemetry[clean_cmd] = data

                        except serial.SerialException as e:
                            logging.error("Error reading from serial device during command %s: %s", clean_cmd, e)
                            cycle_data[clean_cmd] = "Error"
                            raise  # Push up to break context block and trigger re-initialization

                        time.sleep(0.1)

                    # Update timestamps and trigger Excel file updates
                    current_time = time.strftime("%Y-%m-%d %H:%M:%S")
                    live_telemetry["last_updated"] = current_time

                    # Push collected variables block into the layout sheet
                    append_to_excel(current_time, cycle_data)

                    print("Collected data and updated Excel. Going to sleep for 60 seconds...")
                    time.sleep(60)

        except (serial.SerialException, OSError) as init_error:
            logging.error("Serial port connection dropped or missing: %s. Retrying in 10s...", init_error)
            print("Serial port dropped. Re-attempting pipeline in 10 seconds...")
            time.sleep(10)



# ---------------------------------------------------------------------------
# Historical graph + GRMS comparison helpers
# ---------------------------------------------------------------------------

def _parse_timestamp(value, default_date=None):
    """Parse common timestamp formats used by the AMMA-2/GRMS exports."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    # Excel may return a datetime object.
    if isinstance(value, datetime):
        return value

    # ISO and common date/time formats.
    candidates = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%H:%M:%S",
        "%H:%M",
    ]
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        pass

    for fmt in candidates:
        try:
            parsed = datetime.strptime(text, fmt)
            if fmt in ("%H:%M:%S", "%H:%M"):
                if default_date is not None:
                    parsed = datetime.combine(default_date.date(), parsed.time())
            return parsed
        except ValueError:
            continue

    # Numeric Excel serial dates are occasionally exported by spreadsheets.
    try:
        serial = float(text)
        if 20000 < serial < 80000:
            return datetime(1899, 12, 30) + timedelta(days=serial)
    except (ValueError, TypeError):
        pass

    return None


def _numeric_value(value):
    """Extract a numeric value from strings such as '84.2 %'."""
    if value is None:
        return None
    match = re.search(r"-?\d+(?:\.\d+)?", str(value))
    return float(match.group()) if match else None


def _find_header(headers, candidates):
    """Find a column header using case/spacing/punctuation-insensitive matching."""
    normalized = {
        re.sub(r"[^a-z0-9]", "", str(h).lower()): h
        for h in headers
    }
    for candidate in candidates:
        key = re.sub(r"[^a-z0-9]", "", candidate.lower())
        if key in normalized:
            return normalized[key]
    # Fallback: substring matching.
    for h in headers:
        nh = re.sub(r"[^a-z0-9]", "", str(h).lower())
        for candidate in candidates:
            nc = re.sub(r"[^a-z0-9]", "", candidate.lower())
            if nc and (nc in nh or nh in nc):
                return h
    return None


def _filter_comparison_window(rows):
    """Keep one broadcast window: 19:58:00 through 22:50:00.

    If full dates are present, the first date found is treated as the
    broadcast date so that unrelated days are not mixed into the graph.
    """
    if not rows:
        return []

    start_t = datetime.strptime(COMPARISON_START_TIME, "%H:%M:%S").time()
    end_t = datetime.strptime(COMPARISON_END_TIME, "%H:%M:%S").time()

    parsed = []
    for row in rows:
        ts = _parse_timestamp(row.get("timestamp"))
        value = _numeric_value(row.get("value"))
        if ts and value is not None:
            parsed.append((ts, value))

    if not parsed:
        return []

    # Use the first/earliest broadcast date when dates are supplied.
    broadcast_date = min(ts.date() for ts, _ in parsed)
    selected = []
    for ts, value in parsed:
        if ts.date() == broadcast_date and start_t <= ts.time() <= end_t:
            selected.append({
                "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                "value": value
            })
    selected.sort(key=lambda x: x["timestamp"])
    return selected


def _read_amma_upload_rows():
    """Read the uploaded AMMA-2 CSV/XLSX comparison file."""
    if not os.path.exists(AMMA_UPLOAD_FILE):
        return []

    try:
        with open(AMMA_UPLOAD_FILE, "r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            headers = reader.fieldnames or []
            timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
            positive_col = _find_header(
                headers,
                ["Positive Peak (DB)", "Positive Modulation", "Positive Modulation (%)", "DB", "Positive Peak"]
            )
            if not timestamp_col or not positive_col:
                return []

            rows = []
            for item in reader:
                ts = _parse_timestamp(item.get(timestamp_col))
                value = _numeric_value(item.get(positive_col))
                if ts and value is not None:
                    rows.append({
                        "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                        "value": value
                    })
            return _filter_comparison_window(rows)
    except Exception as exc:
        logging.error("Failed reading uploaded AMMA-2 comparison file: %s", exc)
        return []


def read_comparison_amma():
    """Return uploaded AMMA-2 data for the fixed broadcast comparison window."""
    uploaded = _read_amma_upload_rows()
    if uploaded:
        return uploaded

    # Fall back to the live Excel logger if no comparison file has been uploaded.
    return _filter_comparison_window(read_amma_history(24 * 30))


def read_amma_history(hours=24):
    """Read Timestamp + Positive Peak (DB) from the local AMMA Excel log."""
    if not os.path.exists(EXCEL_FILE):
        return []

    cutoff = datetime.now() - timedelta(hours=max(1, min(hours, 24 * 30)))
    rows = []

    try:
        with excel_lock:
            wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
            ws = wb.active
            headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            timestamp_col = _find_header(headers, ["Timestamp"])
            positive_col = _find_header(
                headers,
                ["Positive Peak (DB)", "Positive Modulation", "DB", "Positive Peak"]
            )

            if not timestamp_col or not positive_col:
                wb.close()
                return []

            ts_idx = headers.index(timestamp_col)
            pos_idx = headers.index(positive_col)

            for row in ws.iter_rows(min_row=2, values_only=True):
                ts = _parse_timestamp(row[ts_idx] if ts_idx < len(row) else None)
                value = _numeric_value(row[pos_idx] if pos_idx < len(row) else None)
                if ts and value is not None and ts >= cutoff:
                    rows.append({
                        "timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"),
                        "value": value
                    })
            wb.close()
    except Exception as exc:
        logging.error("Failed reading AMMA historical data: %s", exc)

    return rows


def _read_grms_rows():
    """Read normalized GRMS data and restrict it to the broadcast window."""
    if not os.path.exists(GRMS_FILE):
        return []

    try:
        with open(GRMS_FILE, "r", encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            headers = reader.fieldnames or []
            timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
            positive_col = _find_header(
                headers,
                ["Positive Modulation", "Positive Modulation (%)", "Positive", "Blue", "Positive Peak", "DB"],
            )
            if not timestamp_col or not positive_col:
                logging.error("GRMS file is missing a timestamp or positive-modulation column.")
                return []

            raw_rows = [(item.get(timestamp_col), item.get(positive_col)) for item in reader]

        amma = _read_amma_upload_rows()
        default_date = _parse_timestamp(amma[0]["timestamp"]) if amma else None
        rows = []
        for raw_ts, raw_value in raw_rows:
            ts = _parse_timestamp(raw_ts, default_date=default_date)
            value = _numeric_value(raw_value)
            if ts and value is not None:
                rows.append({"timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"), "value": value})

        return _filter_comparison_window(rows)
    except Exception as exc:
        logging.error("Failed reading GRMS comparison file: %s", exc)
        return []


def _save_grms_upload(file_storage):
    """Normalize an uploaded CSV or XLSX GRMS export into GRMS_FILE."""
    filename = (file_storage.filename or "").lower()

    if filename.endswith(".xlsx"):
        data = file_storage.read()
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        values = list(ws.iter_rows(values_only=True))
        wb.close()

        if not values:
            raise ValueError("The GRMS spreadsheet is empty.")

        headers = [str(v) if v is not None else "" for v in values[0]]
        timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
        positive_col = _find_header(
            headers,
            ["Positive Modulation", "Positive Modulation (%)", "Positive", "Blue", "Positive Peak", "DB"],
        )
        if not timestamp_col or not positive_col:
            raise ValueError(
                "Could not find the GRMS timestamp and positive-modulation columns."
            )

        ts_idx = headers.index(timestamp_col)
        pos_idx = headers.index(positive_col)
        raw_rows = [
            (row[ts_idx] if ts_idx < len(row) else None,
             row[pos_idx] if pos_idx < len(row) else None)
            for row in values[1:]
        ]
    else:
        text = file_storage.read().decode("utf-8-sig")
        sample = text[:4096]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel

        reader = csv.DictReader(io.StringIO(text), dialect=dialect)
        headers = reader.fieldnames or []
        timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
        positive_col = _find_header(
            headers,
            ["Positive Modulation", "Positive Modulation (%)", "Positive", "Blue", "Positive Peak", "DB"],
        )
        if not timestamp_col or not positive_col:
            raise ValueError(
                "Could not find the GRMS timestamp and positive-modulation columns."
            )
        raw_rows = [
            (item.get(timestamp_col), item.get(positive_col))
            for item in reader
        ]

    # Align time-only GRMS exports to the first AMMA logging date.
    amma = read_amma_history(24 * 30)
    default_date = _parse_timestamp(amma[0]["timestamp"]) if amma else datetime.now()

    normalized = []
    for raw_ts, raw_value in raw_rows:
        ts = _parse_timestamp(raw_ts, default_date=default_date)
        value = _numeric_value(raw_value)
        if ts and value is not None:
            normalized.append((ts, value))

    if not normalized:
        raise ValueError("No usable timestamp/positive-modulation rows were found.")

    normalized.sort(key=lambda item: item[0])

    with open(GRMS_FILE, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Timestamp", "Positive Modulation"])
        for ts, value in normalized:
            writer.writerow([ts.strftime("%Y-%m-%d %H:%M:%S"), value])

    return len(normalized)


def _save_amma_upload(file_storage):
    """Normalize an uploaded AMMA-2 CSV/XLSX export into AMMA_UPLOAD_FILE."""
    filename = (file_storage.filename or "").lower()

    if filename.endswith(".xlsx"):
        data = file_storage.read()
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        values = list(ws.iter_rows(values_only=True))
        wb.close()
        if not values:
            raise ValueError("The AMMA-2 spreadsheet is empty.")

        headers = [str(v) if v is not None else "" for v in values[0]]
        timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
        positive_col = _find_header(headers, [
            "Positive Peak (DB)", "Positive Modulation", "Positive Modulation (%)", "DB", "Positive Peak"
        ])
        if not timestamp_col or not positive_col:
            raise ValueError("Could not find the AMMA-2 Timestamp and DB Positive Modulation columns.")
        ti, pi = headers.index(timestamp_col), headers.index(positive_col)
        raw_rows = [(row[ti] if ti < len(row) else None, row[pi] if pi < len(row) else None) for row in values[1:]]
    else:
        text = file_storage.read().decode("utf-8-sig")
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(io.StringIO(text), dialect=dialect)
        headers = reader.fieldnames or []
        timestamp_col = _find_header(headers, ["Timestamp", "Date Time", "DateTime", "Time", "Date"])
        positive_col = _find_header(headers, [
            "Positive Peak (DB)", "Positive Modulation", "Positive Modulation (%)", "DB", "Positive Peak"
        ])
        if not timestamp_col or not positive_col:
            raise ValueError("Could not find the AMMA-2 Timestamp and DB Positive Modulation columns.")
        raw_rows = [(item.get(timestamp_col), item.get(positive_col)) for item in reader]

    normalized = []
    for raw_ts, raw_value in raw_rows:
        ts = _parse_timestamp(raw_ts)
        value = _numeric_value(raw_value)
        if ts and value is not None:
            normalized.append((ts, value))

    if not normalized:
        raise ValueError("No usable AMMA-2 timestamp/DB rows were found.")

    normalized.sort(key=lambda x: x[0])
    with open(AMMA_UPLOAD_FILE, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Timestamp", "Positive Modulation"])
        for ts, value in normalized:
            writer.writerow([ts.strftime("%Y-%m-%d %H:%M:%S"), value])

    windowed = _filter_comparison_window([
        {"timestamp": ts.strftime("%Y-%m-%dT%H:%M:%S"), "value": value}
        for ts, value in normalized
    ])
    if not windowed:
        raise ValueError("The AMMA-2 file contains no readings between 19:58:00 and 22:50:00.")
    return len(windowed)


def calculate_comparison(hours=24):
    """Match AMMA DB readings to the nearest GRMS positive-modulation point."""
    amma = read_comparison_amma()
    grms = _read_grms_rows()

    if not amma or not grms:
        return {
            "matched_points": 0,
            "mean_absolute_difference": None,
            "correlation": None,
        }

    amma_pairs = [
        (_parse_timestamp(row["timestamp"]), row["value"])
        for row in amma
    ]
    grms_pairs = [
        (_parse_timestamp(row["timestamp"]), row["value"])
        for row in grms
    ]

    grms_pairs = [(ts, val) for ts, val in grms_pairs if ts is not None]
    grms_pairs.sort(key=lambda item: item[0])
    grms_times = [item[0] for item in grms_pairs]

    matched_a = []
    matched_g = []

    max_gap = timedelta(seconds=90)
    for a_ts, a_val in amma_pairs:
        idx = bisect_left(grms_times, a_ts)
        candidates = []
        if idx < len(grms_pairs):
            candidates.append(grms_pairs[idx])
        if idx > 0:
            candidates.append(grms_pairs[idx - 1])
        if not candidates:
            continue

        g_ts, g_val = min(candidates, key=lambda item: abs(item[0] - a_ts))
        if abs(g_ts - a_ts) <= max_gap:
            matched_a.append(a_val)
            matched_g.append(g_val)

    if not matched_a:
        return {
            "matched_points": 0,
            "mean_absolute_difference": None,
            "correlation": None,
        }

    mae = sum(abs(a - g) for a, g in zip(matched_a, matched_g)) / len(matched_a)

    correlation = None
    if len(matched_a) >= 2:
        mean_a = sum(matched_a) / len(matched_a)
        mean_g = sum(matched_g) / len(matched_g)
        num = sum((a - mean_a) * (g - mean_g) for a, g in zip(matched_a, matched_g))
        den_a = math.sqrt(sum((a - mean_a) ** 2 for a in matched_a))
        den_g = math.sqrt(sum((g - mean_g) ** 2 for g in matched_g))
        if den_a and den_g:
            correlation = num / (den_a * den_g)

    return {
        "matched_points": len(matched_a),
        "mean_absolute_difference": round(mae, 2),
        "correlation": round(correlation, 3) if correlation is not None else None,
    }


# HTML Interface embedded directly for a single-file portable solution
DASHBOARD_HTML: str = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TWR Africa — Belar AMMA-2 Real-Time Modulation Monitor</title>
    <style>
        :root {
            --bg-main: #0b0f19;
            --bg-card: #111827;
            --border-color: #1f2937;
            --text-main: #f3f4f6;
            --text-muted: #9ca3af;
            --twr-blue: #00a3e0;
            --twr-navy: #0f2c59;
            --signal-ok: #10b981;
            --signal-warn: #f59e0b;
            --signal-crit: #ef4444;
            --led-off: #1f2937;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            background-color: var(--bg-main);
            color: var(--text-main);
            min-height: 100vh;
            padding: 24px;
        }

        .container {
            max-width: 1000px;
            margin: 0 auto;
        }

        /* Broadcast Console Rack-Mount Header Layout */
        .header {
            background: linear-gradient(90deg, var(--twr-navy) 0%, #071630 100%);
            padding: 20px 24px;
            border-radius: 6px 6px 0 0;
            border-bottom: 3px solid var(--twr-blue);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
        }

        .header-title h1 {
            font-size: 1.4rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            color: #ffffff;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .header-title h2 {
            color: var(--twr-blue);
            font-size: 0.85rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 2px;
        }

        .system-badge {
            background: rgba(0, 163, 224, 0.15);
            border: 1px solid var(--twr-blue);
            color: #ffffff;
            font-size: 0.7rem;
            padding: 4px 10px;
            border-radius: 4px;
            font-weight: 700;
            letter-spacing: 0.5px;
        }

        /* Master Control Status Bar */
        .status-bar {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-top: none;
            padding: 12px 24px;
            border-radius: 0 0 6px 6px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            font-size: 0.85rem;
        }

        .status-indicator {
            display: flex;
            align-items: center;
            gap: 10px;
            font-weight: 600;
            letter-spacing: 0.5px;
        }

        .pulse-dot {
            width: 10px;
            height: 10px;
            background-color: var(--signal-ok);
            border-radius: 50%;
            animation: telemetryPulse 1.5s ease-in-out infinite;
            box-shadow: 0 0 8px var(--signal-ok);
        }

        @keyframes telemetryPulse {
            0%, 100% { opacity: 1; transform: scale(1); }
            50% { opacity: 0.4; transform: scale(1.1); }
        }

        .refresh-controls {
            display: flex;
            align-items: center;
            gap: 16px;
            color: var(--text-muted);
        }

        .refresh-btn {
            background: #1f2937;
            border: 1px solid #374151;
            color: var(--text-main);
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .refresh-btn:hover {
            background: #374151;
            border-color: #4b5563;
        }

        /* High-Density Signal Monitoring Grid Layout */
        .telemetry-grid {
            display: flex;
            flex-direction: column;
            gap: 12px;
            margin-bottom: 20px;
        }

        .telemetry-card {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 16px 20px;
            display: grid;
            grid-template-columns: 280px 1fr 140px;
            align-items: center;
            gap: 24px;
            position: relative;
        }

        /* Layout fallback adjustments for ancillary metrics without hardware meters */
        .telemetry-card.no-meter {
            grid-template-columns: 280px 1fr 140px;
        }

        .telemetry-card.no-meter .meter-wrapper {
            display: flex;
            align-items: center;
            color: #4b5563;
            font-size: 0.8rem;
            font-style: italic;
            letter-spacing: 0.5px;
        }

        .metric-details {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .metric-tag {
            font-size: 0.85rem;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: 0.5px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .metric-tag span {
            font-size: 0.7rem;
            background: #1f2937;
            padding: 1px 6px;
            border-radius: 3px;
            color: var(--twr-blue);
        }

        .metric-desc {
            font-size: 0.75rem;
            color: var(--text-muted);
        }

        /* Hardware LED Segment Display Engineering Emulator */
        .meter-wrapper {
            width: 100%;
        }

        .led-strip {
            display: flex;
            gap: 3px;
            background: #020617;
            padding: 6px;
            border-radius: 4px;
            border: 1px solid #1e293b;
        }

        .led-segment {
            flex: 1;
            height: 14px;
            background-color: var(--led-off);
            border-radius: 1px;
            transition: background-color 0.1s ease;
        }

        /* Master Scale Labels below the meter segments */
        .scale-labels {
            display: flex;
            justify-content: space-between;
            margin-top: 4px;
            font-family: 'Courier New', monospace;
            font-size: 0.65rem;
            color: #4b5563;
            font-weight: 700;
            padding: 0 2px;
        }

        /* Digital Nixie Tube Display Box */
        .readout-container {
            text-align: right;
        }

        .digital-readout {
            font-family: 'Courier New', monospace;
            font-size: 1.7rem;
            font-weight: 700;
            color: var(--signal-ok);
            background-color: #050b14;
            border: 1px solid #1e293b;
            padding: 6px 12px;
            border-radius: 4px;
            display: inline-block;
            min-width: 120px;
            text-align: center;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.8);
            letter-spacing: -0.5px;
        }

        /* Specific Status Variations based on RF compliance math thresholds */
        .telemetry-card.alert-flash .digital-readout {
            color: var(--signal-crit) !important;
            animation: criticalFlash 0.5s ease-in-out infinite alternate;
        }

        @keyframes criticalFlash {
            0% { box-shadow: inset 0 2px 4px rgba(0,0,0,0.8); }
            100% { box-shadow: inset 0 0 12px rgba(239, 68, 68, 0.4); border-color: var(--signal-crit); }
        }

        .telemetry-card.warn-state .digital-readout {
            color: var(--signal-warn) !important;
        }

        /* Monitoring Station System Footer Configuration */
        .footer {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 16px 24px;
            font-size: 0.8rem;
        }

        .sync-time {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 12px;
            color: var(--text-main);
        }

        .sync-time strong {
            color: var(--twr-blue);
            font-family: 'Courier New', monospace;
        }

        .metadata-line {
            display: flex;
            gap: 24px;
            flex-wrap: wrap;
            border-top: 1px solid var(--border-color);
            padding-top: 12px;
            color: var(--text-muted);
        }

        .metadata-line span strong {
            color: var(--text-main);
        }

        /* Responsive Breakpoints for Field Inspection Terminals */
        @media (max-width: 850px) {
            .telemetry-card {
                grid-template-columns: 1fr;
                gap: 12px;
            }
            .readout-container {
                text-align: left;
            }
            .digital-readout {
                width: 100%;
            }
        }

        /* Live Data Transit Interception Flash */
        .data-pulse {
            animation: visualCommit 0.3s ease-out;
        }

        @keyframes visualCommit {
            0% { opacity: 0.7; transform: scale(0.995); }
            100% { opacity: 1; transform: scale(1); }
        }

        /* Historical graph / comparison panel */
        .history-panel {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 6px;
            padding: 20px 24px;
            margin-bottom: 20px;
        }

        .history-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 20px;
            flex-wrap: wrap;
            margin-bottom: 16px;
        }

        .history-header h2 {
            font-size: 1.05rem;
            margin-bottom: 5px;
            color: #ffffff;
        }

        .history-header p {
            color: var(--text-muted);
            font-size: 0.78rem;
        }

        .history-controls {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
        }

        .history-controls label,
        .upload-label {
            color: var(--text-muted);
            font-size: 0.75rem;
            font-weight: 600;
        }

        .history-controls select,
        .upload-row input[type="file"] {
            background: #020617;
            color: var(--text-main);
            border: 1px solid #374151;
            border-radius: 4px;
            padding: 7px 9px;
            font-size: 0.75rem;
        }

        .upload-row {
            border-top: 1px solid var(--border-color);
            border-bottom: 1px solid var(--border-color);
            padding: 12px 0;
            margin-bottom: 16px;
        }

        #grmsUploadForm {
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }

        .upload-btn {
            background: var(--twr-navy);
            border: 1px solid var(--twr-blue);
            color: #ffffff;
            padding: 7px 12px;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 700;
            cursor: pointer;
        }

        .upload-btn:hover {
            background: #123b72;
        }

        #uploadStatus {
            color: var(--text-muted);
            font-size: 0.72rem;
        }

        .comparison-stats {
            display: grid;
            grid-template-columns: repeat(5, minmax(120px, 1fr));
            gap: 8px;
            margin-bottom: 14px;
        }

        .stat-box {
            background: #050b14;
            border: 1px solid #1e293b;
            border-radius: 4px;
            padding: 10px;
        }

        .stat-box span {
            display: block;
            color: var(--text-muted);
            font-size: 0.65rem;
            margin-bottom: 5px;
        }

        .stat-box strong {
            color: var(--twr-blue);
            font-family: 'Courier New', monospace;
            font-size: 1rem;
        }

        .chart-wrapper {
            position: relative;
            width: 100%;
            height: 430px;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            overflow: hidden;
        }

        #modulationChart {
            display: block;
            width: 100%;
            height: 100%;
            cursor: crosshair;
            touch-action: pan-y;
        }

        .chart-note {
            font-size: 0.72rem;
            color: #94a3b8;
            margin-top: 6px;
        }

        .series-summary {
            display: grid;
            grid-template-columns: 1fr;
            gap: 6px;
            margin-top: 12px;
        }

        .series-summary-item {
            background: #050b14;
            border: 1px solid #1e293b;
            border-radius: 4px;
            padding: 8px 12px;
            font-size: 0.74rem;
            color: var(--text-muted);
            display: flex;
            gap: 14px;
            flex-wrap: wrap;
            align-items: baseline;
        }

        .series-summary-item strong {
            color: var(--text-main);
            min-width: 110px;
        }

        .series-summary-item span {
            font-family: 'Courier New', monospace;
            color: var(--twr-blue);
        }

        .chart-message {
            position: absolute;
            inset: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            color: #475569;
            font-size: 0.8rem;
            pointer-events: none;
        }

        .chart-legend {
            display: flex;
            gap: 22px;
            flex-wrap: wrap;
            margin-top: 10px;
            color: var(--text-muted);
            font-size: 0.72rem;
        }

        .legend-dot {
            display: inline-block;
            width: 9px;
            height: 9px;
            border-radius: 50%;
            margin-right: 5px;
        }

        .amma-dot {
            background: #2563eb;
        }

        .grms-dot {
            background: #ef4444;
        }

        .grms-help {
            margin-top: 12px;
            color: var(--text-muted);
            font-size: 0.7rem;
            line-height: 1.5;
        }

        .grms-help code {
            color: #ffffff;
            background: #111827;
            padding: 2px 5px;
            border-radius: 3px;
        }

        @media (max-width: 900px) {
            .comparison-stats {
                grid-template-columns: repeat(2, 1fr);
            }
        }

    </style>
</head>
<body>
    <div class="container">
        <!-- Rack Frame Top Header -->
        <div class="header">
            <div class="header-title">
                <h1>📡 TWR AFRICA — AM MONITORING BRIDGE</h1>
                <h2>Belar AMMA-2 Precision Receiver Node</h2>
            </div>
            <div class="system-badge">LINK: TELEMETRY DOWNLINK</div>
        </div>

        <!-- System Diagnostic Line -->
        <div class="status-bar">
            <div class="status-indicator">
                <div class="pulse-dot" id="sysPulse"></div>
                <span id="sysStatusText">CONSOLE STREAMING ACTIVE</span>
            </div>
            <div class="refresh-controls">
                <span>CADENCE: <strong id="refreshCounter" style="color: var(--text-main);">5</strong>s</span>
                <button class="refresh-btn" onclick="toggleAutoRefresh()">⏸ PAUSE BUS</button>
            </div>
        </div>

        <!-- Operational Instruments Matrix Stack -->
        <div class="telemetry-grid">

            <!-- INSTRUMENT 1: NEGATIVE PEAK (DA) -->
            <div class="telemetry-card" id="cardDA">
                <div class="metric-details">
                    <div class="metric-tag">NEGATIVE PEAK MODULATION <span>DA</span></div>
                    <div class="metric-desc">Carrier dropout evaluation (Limit: 100% standard envelope safety)</div>
                </div>
                <div class="meter-wrapper">
                    <div class="led-strip" id="stripDA"></div>
                    <div class="scale-labels">
                        <span>0%</span><span>20</span><span>40</span><span>60</span><span>80</span><span>100%</span><span>120</span><span>140%</span>
                    </div>
                </div>
                <div class="readout-container">
                    <div class="digital-readout" id="valueDA">{{ telemetry.DA }}</div>
                </div>
            </div>

            <!-- INSTRUMENT 2: POSITIVE PEAK (DB) -->
            <div class="telemetry-card" id="cardDB">
                <div class="metric-details">
                    <div class="metric-tag">POSITIVE PEAK MODULATION <span>DB</span></div>
                    <div class="metric-desc">Asymmetric envelope modulation crest factor (Max allocation: 125%)</div>
                </div>
                <div class="meter-wrapper">
                    <div class="led-strip" id="stripDB"></div>
                    <div class="scale-labels">
                        <span>0%</span><span>20</span><span>40</span><span>60</span><span>80</span><span>100%</span><span>120</span><span>140%</span>
                    </div>
                </div>
                <div class="readout-container">
                    <div class="digital-readout" id="valueDB">{{ telemetry.DB }}</div>
                </div>
            </div>

            <!-- INSTRUMENT 3: CARRIER REFERENCE (DG) -->
            <div class="telemetry-card no-meter" id="cardDG">
                <div class="metric-details">
                    <div class="metric-tag">UNMODULATED CARRIER REFERENCE <span>DG</span></div>
                    <div class="metric-desc">RF power engine normalization calibration level</div>
                </div>
                <div class="meter-wrapper">
                    <div>[ Dynamic Linear Reference Signal - Static Variable Tracking ]</div>
                </div>
                <div class="readout-container">
                    <div class="digital-readout" id="valueDG" style="color: #38bdf8;">{{ telemetry.DG }}</div>
                </div>
            </div>

            <!-- INSTRUMENT 4: AM NOISE FLOOR (DH) -->
            <div class="telemetry-card no-meter" id="cardDH">
                <div class="metric-details">
                    <div class="metric-tag">RESIDUAL AM NOISE FLOOR <span>DH</span></div>
                    <div class="metric-desc">Integrated thermal noise &amp; residual interference (Target &lt; -50dB)</div>
                </div>
                <div class="meter-wrapper">
                    <div>[ Spectrum Isolation Subsystem - Decibel Logarithmic Value ]</div>
                </div>
                <div class="readout-container">
                    <div class="digital-readout" id="valueDH">{{ telemetry.DH }}</div>
                </div>
            </div>

            <!-- INSTRUMENT 5: RMS MODULATION (DN) -->
            <div class="telemetry-card" id="cardDN">
                <div class="metric-details">
                    <div class="metric-tag">AVERAGE RMS MODULATION DEPTH <span>DN</span></div>
                    <div class="metric-desc">Integrated root-mean-square continuous audio energy levels</div>
                </div>
                <div class="meter-wrapper">
                    <div class="led-strip" id="stripDN"></div>
                    <div class="scale-labels">
                        <span>0%</span><span>20</span><span>40</span><span>60</span><span>80</span><span>100%</span><span>120%</span>
                    </div>
                </div>
                <div class="readout-container">
                    <div class="digital-readout" id="valueDN">{{ telemetry.DN }}</div>
                </div>
            </div>

        </div>


        <!-- Historical Positive Modulation / GRMS Comparison -->
        <div class="history-panel">
            <div class="history-header">
                <div>
                    <h2>Positive Modulation History &amp; GRMS Comparison</h2>
                    <p>Upload the real AMMA-2 broadcast data and GRMS export to compare Positive Modulation (DB) over the same broadcast window.</p>
                </div>
                <div class="history-controls">
                    <label>Broadcast window</label>
                    <strong style="color:var(--text-main);">19:58:00 → 22:50:00</strong>
                    <button class="refresh-btn" onclick="refreshHistory()">↻ REFRESH GRAPH</button>
                </div>
            </div>

            <div class="upload-row" style="display:flex;gap:24px;flex-wrap:wrap;">
                <form id="ammaUploadForm">
                    <label class="upload-label" for="ammaFile">AMMA-2 CSV/XLSX</label>
                    <input id="ammaFile" name="file" type="file" accept=".csv,.xlsx" required>
                    <button type="submit" class="upload-btn">UPLOAD AMMA-2 DATA</button>
                    <span id="ammaUploadStatus">No AMMA-2 comparison file loaded.</span>
                </form>
                <form id="grmsUploadForm">
                    <label class="upload-label" for="grmsFile">GRMS CSV/XLSX</label>
                    <input id="grmsFile" name="file" type="file" accept=".csv,.xlsx" required>
                    <button type="submit" class="upload-btn">UPLOAD GRMS DATA</button>
                    <span id="uploadStatus">No GRMS file loaded.</span>
                </form>
            </div>

            <div class="comparison-stats">
                <div class="stat-box">
                    <span>AMMA points</span>
                    <strong id="ammaCount">0</strong>
                </div>
                <div class="stat-box">
                    <span>GRMS points</span>
                    <strong id="grmsCount">0</strong>
                </div>
                <div class="stat-box">
                    <span>Matched points</span>
                    <strong id="matchedCount">0</strong>
                </div>
                <div class="stat-box">
                    <span>Mean absolute difference</span>
                    <strong id="maeValue">—</strong>
                </div>
                <div class="stat-box">
                    <span>Correlation</span>
                    <strong id="correlationValue">—</strong>
                </div>
            </div>

            <div class="chart-wrapper">
                <canvas id="modulationChart"></canvas>
                <div id="chartMessage" class="chart-message">Upload both files to compare the broadcast.</div>
            </div>
            <div class="chart-note">Comparison window: {{ window_start }} → {{ window_end }} · hover the graph for exact readings</div>

            <div class="series-summary" id="seriesSummary"></div>

            <div class="chart-legend">
                <span><i class="legend-dot amma-dot"></i> AMMA-2 Positive Modulation (DB) — blue</span>
                <span><i class="legend-dot grms-dot"></i> GRMS Positive Modulation — red</span>
            </div>

            <div class="grms-help">
                <strong>GRMS file format:</strong> CSV or XLSX containing a timestamp/time column and a
                positive-modulation column. Example:
                <code>Timestamp,Positive Modulation</code>
            </div>
        </div>

        <!-- Infrastructure Meta Blocks -->
        <div class="footer">
            <div class="sync-time">
                ⏱ Master Bus Sync: <strong id="lastUpdated">{{ telemetry.last_updated }}</strong>
            </div>
            <div class="metadata-line">
                <span>STORAGE SYSTEM: <strong>data.log</strong> &amp; <strong>modulation_data.xlsx</strong></span>
                <span>DATA RATE: <strong>9600 bps (8N1)</strong></span>
                <span>BUS TIMEOUT: <strong>5000 ms</strong></span>
            </div>
        </div>
    </div>

    <script>
        let autoRefresh = true;
        let refreshInterval;
        let countdownInterval;
        const SEGMENT_COUNT = 40; // High-density resolution meter segments

        // Generate Hardware-Style LED segments programmatically
        function structuralBuildLeds() {
            const targets = ['stripDA', 'stripDB', 'stripDN'];
            targets.forEach(id => {
                const container = document.getElementById(id);
                if (!container) return;
                container.innerHTML = '';
                for (let i = 0; i < SEGMENT_COUNT; i++) {
                    const seg = document.createElement('div');
                    seg.className = 'led-segment';
                    container.appendChild(seg);
                }
            });
        }

        function extractNumericVolume(valStr) {
            if (!valStr || valStr.includes("No Data") || valStr.includes("Error")) return null;
            const outMatch = valStr.match(/(-?[0-9]+(?:[.][0-9]+)?)/);
            return outMatch ? parseFloat(outMatch[1]) : null;
        }

        // Render standard color-coded broadcast segments (Green -> Orange -> Crimson)
        function driveHardwareMeter(stripId, percentageValue, warnThreshold, critThreshold) {
            const strip = document.getElementById(stripId);
            if (!strip) return;
            const segments = strip.getElementsByClassName('led-segment');
            const activeCount = Math.round((Math.min(Math.max(percentageValue, 0), 150) / 150) * SEGMENT_COUNT);

            for (let i = 0; i < SEGMENT_COUNT; i++) {
                if (i < activeCount) {
                    const currentPct = (i / SEGMENT_COUNT) * 150;
                    if (currentPct >= critThreshold) {
                        segments[i].style.backgroundColor = 'var(--signal-crit)';
                        segments[i].style.boxShadow = '0 0 6px var(--signal-crit)';
                    } else if (currentPct >= warnThreshold) {
                        segments[i].style.backgroundColor = 'var(--signal-warn)';
                        segments[i].style.boxShadow = '0 0 4px var(--signal-warn)';
                    } else {
                        segments[i].style.backgroundColor = 'var(--signal-ok)';
                        segments[i].style.boxShadow = 'none';
                    }
                } else {
                    segments[i].style.backgroundColor = 'var(--led-off)';
                    segments[i].style.boxShadow = 'none';
                }
            }
        }

        function evaluateInstrumentCompliance() {
            // Processing Instrument DA (Negative Peak)
            const daVal = extractNumericVolume(document.getElementById('valueDA').innerText);
            const cardDA = document.getElementById('cardDA');
            if (daVal !== null) {
                driveHardwareMeter('stripDA', daVal, 85, 99); // Critical danger at 100% negative (carrier cut)
                cardDA.className = 'telemetry-card' + (daVal >= 99 ? ' alert-flash' : daVal >= 85 ? ' warn-state' : '');
            }

            // Processing Instrument DB (Positive Peak)
            const dbVal = extractNumericVolume(document.getElementById('valueDB').innerText);
            const cardDB = document.getElementById('cardDB');
            if (dbVal !== null) {
                driveHardwareMeter('stripDB', dbVal, 100, 125); // Asymmetrical peaks headroom standard
                cardDB.className = 'telemetry-card' + (dbVal >= 125 ? ' alert-flash' : dbVal >= 100 ? ' warn-state' : '');
            }

            // Processing Instrument DH (Noise Floor evaluation in Decibels)
            const dhVal = extractNumericVolume(document.getElementById('valueDH').innerText);
            const cardDH = document.getElementById('cardDH');
            if (dhVal !== null) {
                // For decibels, values closer to 0 (e.g. -40dB) represent higher noise than -60dB.
                cardDH.className = 'telemetry-card no-meter' + (dhVal >= -45 ? ' alert-flash' : dhVal >= -50 ? ' warn-state' : '');
                document.getElementById('valueDH').style.color = dhVal >= -45 ? 'var(--signal-crit)' : dhVal >= -50 ? 'var(--signal-warn)' : 'var(--signal-ok)';
            }

            // Processing Instrument DN (RMS Audio Tracking)
            const dnVal = extractNumericVolume(document.getElementById('valueDN').innerText);
            const cardDN = document.getElementById('cardDN');
            if (dnVal !== null) {
                // Adjusting scale calculation matrix to target 125% volume mapping reference
                const normalizedDN = (dnVal / 125) * 150;
                driveHardwareMeter('stripDN', normalizedDN, 85, 110);
                cardDN.className = 'telemetry-card' + (dnVal >= 110 ? ' alert-flash' : dnVal >= 85 ? ' warn-state' : '');
            }
        }

        function updateDashboard() {
            fetch('/api/telemetry')
                .then(res => res.json())
                .then(data => {
                    document.getElementById('sysPulse').style.backgroundColor = 'var(--signal-ok)';
                    document.getElementById('sysPulse').style.boxShadow = '0 0 8px var(--signal-ok)';
                    document.getElementById('sysStatusText').innerText = "CONSOLE STREAMING ACTIVE";

                    const metrics = ['DA', 'DB', 'DG', 'DH', 'DN'];
                    metrics.forEach(m => {
                        const cell = document.getElementById(`value${m}`);
                        const incoming = data[m] || "No Data";
                        if (cell.innerText !== incoming) {
                            cell.innerText = incoming;
                            const card = document.getElementById(`card${m}`);
                            card.classList.add('data-pulse');
                            setTimeout(() => card.classList.remove('data-pulse'), 300);
                        }
                    });

                    document.getElementById('lastUpdated').innerText = data.last_updated || "Never";
                    evaluateInstrumentCompliance();
                })
                .catch(err => {
                    console.error('Bus Downlink Exception:', err);
                    document.getElementById('sysPulse').style.backgroundColor = 'var(--signal-crit)';
                    document.getElementById('sysPulse').style.boxShadow = '0 0 8px var(--signal-crit)';
                    document.getElementById('sysStatusText').innerText = "LINK TIMEOUT / HARDWARE FAULT";
                });
        }

        function startAutoRefresh() {
            if (refreshInterval) clearInterval(refreshInterval);
            if (countdownInterval) clearInterval(countdownInterval);

            let sec = 5;
            const counter = document.getElementById('refreshCounter');
            if (counter) counter.innerText = sec;

            refreshInterval = setInterval(() => {
                if (autoRefresh) {
                    updateDashboard();
                    sec = 5;
                    if (counter) counter.innerText = sec;
                }
            }, 5000);

            countdownInterval = setInterval(() => {
                if (autoRefresh) {
                    sec--;
                    if (counter && sec >= 0) counter.innerText = sec;
                    if (sec < 0) sec = 4;
                }
            }, 1000);
        }

        function toggleAutoRefresh() {
            autoRefresh = !autoRefresh;
            const btn = document.querySelector('.refresh-btn');
            if (autoRefresh) {
                btn.innerHTML = '⏸ PAUSE BUS';
                btn.style.background = '#1f2937';
                updateDashboard();
                startAutoRefresh();
            } else {
                btn.innerHTML = '▶ RESUME BUS';
                btn.style.background = 'var(--twr-navy)';
                if (refreshInterval) clearInterval(refreshInterval);
                if (countdownInterval) clearInterval(countdownInterval);
            }
        }


        // -------------------------------------------------------------------
        // Historical Positive Modulation / GRMS comparison graph
        // -------------------------------------------------------------------

        function getCanvasContext() {
            const canvas = document.getElementById('modulationChart');
            if (!canvas) return null;
            const rect = canvas.getBoundingClientRect();
            const dpr = window.devicePixelRatio || 1;
            const width = Math.max(300, Math.floor(rect.width));
            const height = Math.max(250, Math.floor(rect.height));

            if (canvas.width !== width * dpr || canvas.height !== height * dpr) {
                canvas.width = width * dpr;
                canvas.height = height * dpr;
            }

            const ctx = canvas.getContext('2d');
            ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
            return { ctx, width, height };
        }

        // Broadcast window comes from the server so the Python constants stay the single source of truth.
        const WINDOW_START = "{{ window_start }}";
        const WINDOW_END = "{{ window_end }}";

        // Latest chart data and hover position, kept so resize/hover redraw without refetching.
        let chartState = null;
        let chartHoverX = null;
        let lastChartRows = { amma: [], grms: [] };

        function chartDate(value) {
            const d = new Date(value);
            return Number.isNaN(d.getTime()) ? null : d;
        }

        function pad2(n) {
            return String(n).padStart(2, '0');
        }

        function formatClock(ms, withSeconds) {
            const d = new Date(ms);
            return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) +
                (withSeconds ? ':' + pad2(d.getSeconds()) : '');
        }

        function formatChartTime(date) {
            return formatClock(date.getTime(), false);
        }

        // Broadcast window (start/end in ms) on the same calendar day as dayMs.
        function windowBounds(dayMs) {
            const day = new Date(dayMs);
            const at = hms => {
                const parts = hms.split(':').map(Number);
                return new Date(day.getFullYear(), day.getMonth(), day.getDate(),
                                parts[0] || 0, parts[1] || 0, parts[2] || 0).getTime();
            };
            return { start: at(WINDOW_START), end: at(WINDOW_END) };
        }

        // Pick a readable gridline spacing for the visible time span.
        function pickTickStep(spanMs) {
            const minute = 60000;
            if (spanMs <= 15 * minute) return minute;
            if (spanMs <= 60 * minute) return 5 * minute;
            if (spanMs <= 3 * 60 * minute) return 15 * minute;
            if (spanMs <= 6 * 60 * minute) return 30 * minute;
            return 60 * minute;
        }

        function medianInterval(points) {
            if (points.length < 2) return 0;
            const diffs = [];
            for (let i = 1; i < points.length; i++) diffs.push(points[i].time - points[i - 1].time);
            diffs.sort((a, b) => a - b);
            return diffs[Math.floor(diffs.length / 2)];
        }

        function summarizeSeries(points) {
            if (!points.length) return null;
            let min = Infinity, max = -Infinity, sum = 0, over100 = 0, over125 = 0, maxAt = points[0].time;
            points.forEach(p => {
                if (p.value < min) min = p.value;
                if (p.value > max) { max = p.value; maxAt = p.time; }
                sum += p.value;
                if (p.value >= 100) over100++;
                if (p.value >= 125) over125++;
            });
            return { count: points.length, min, max, avg: sum / points.length, over100, over125, maxAt };
        }

        function nearestPoint(points, t) {
            if (!points.length) return null;
            let lo = 0, hi = points.length - 1;
            while (hi - lo > 1) {
                const mid = (lo + hi) >> 1;
                if (points[mid].time < t) lo = mid; else hi = mid;
            }
            return Math.abs(points[lo].time - t) <= Math.abs(points[hi].time - t) ? points[lo] : points[hi];
        }

        // Draw one series as a line that breaks across logging gaps, plus point markers.
        function drawSeries(ctx, points, scale, plot, strokeStyle) {
            if (!points.length) return;
            const gapMs = Math.max(150000, medianInterval(points) * 3);
            ctx.save();
            ctx.beginPath();
            ctx.rect(plot.left, plot.top, plot.width, plot.height);
            ctx.clip();

            ctx.strokeStyle = strokeStyle;
            ctx.lineWidth = 1.6;
            ctx.lineJoin = 'round';
            ctx.lineCap = 'round';
            ctx.beginPath();
            let prev = null;
            points.forEach(point => {
                const px = scale.x(point.time);
                const py = scale.y(point.value);
                if (prev && point.time - prev.time <= gapMs) ctx.lineTo(px, py);
                else ctx.moveTo(px, py);
                prev = point;
            });
            ctx.stroke();

            // Markers show real sample positions (and isolated samples either side of a gap).
            if (points.length <= 600) {
                ctx.fillStyle = strokeStyle;
                points.forEach(point => {
                    ctx.beginPath();
                    ctx.arc(scale.x(point.time), scale.y(point.value), 2.2, 0, Math.PI * 2);
                    ctx.fill();
                });
            }
            ctx.restore();
        }

        function drawLegend(ctx, plot, amma, grms) {
            const items = [
                { color: '#2563eb', text: 'AMMA-2 Positive Modulation (DB)  n=' + amma.length },
                { color: '#dc2626', text: 'GRMS Positive Modulation  n=' + grms.length }
            ];
            ctx.save();
            ctx.font = '11px Segoe UI, Arial, sans-serif';
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            const boxW = Math.max(...items.map(i => ctx.measureText(i.text).width)) + 34;
            const boxH = items.length * 18 + 8;
            const x = plot.left + 10;
            const y = plot.top + 8;
            ctx.fillStyle = 'rgba(255,255,255,0.92)';
            ctx.strokeStyle = '#cbd5e1';
            ctx.lineWidth = 1;
            ctx.fillRect(x, y, boxW, boxH);
            ctx.strokeRect(x + 0.5, y + 0.5, boxW, boxH);
            items.forEach((item, i) => {
                const cy = y + 4 + i * 18 + 9;
                ctx.strokeStyle = item.color;
                ctx.lineWidth = 2;
                ctx.beginPath(); ctx.moveTo(x + 8, cy); ctx.lineTo(x + 22, cy); ctx.stroke();
                ctx.fillStyle = item.color;
                ctx.beginPath(); ctx.arc(x + 15, cy, 2.4, 0, Math.PI * 2); ctx.fill();
                ctx.fillStyle = '#334155';
                ctx.fillText(item.text, x + 28, cy);
            });
            ctx.restore();
        }

        // Crosshair + tooltip with the nearest reading from each series.
        function drawHover(ctx, plot, scale, state, width) {
            if (chartHoverX === null || chartHoverX < plot.left || chartHoverX > plot.left + plot.width) return;
            const t = state.xMin + ((chartHoverX - plot.left) / plot.width) * (state.xMax - state.xMin);
            const tolerance = 2 * 60000;
            const a = nearestPoint(state.amma, t);
            const g = nearestPoint(state.grms, t);
            const aHit = a && Math.abs(a.time - t) <= tolerance ? a : null;
            const gHit = g && Math.abs(g.time - t) <= tolerance ? g : null;

            ctx.save();
            ctx.strokeStyle = 'rgba(15,23,42,0.45)';
            ctx.lineWidth = 1;
            ctx.setLineDash([3, 3]);
            ctx.beginPath(); ctx.moveTo(chartHoverX, plot.top); ctx.lineTo(chartHoverX, plot.top + plot.height); ctx.stroke();
            ctx.setLineDash([]);

            [[aHit, '#2563eb'], [gHit, '#dc2626']].forEach(entry => {
                if (!entry[0]) return;
                ctx.fillStyle = '#ffffff';
                ctx.strokeStyle = entry[1];
                ctx.lineWidth = 2;
                ctx.beginPath();
                ctx.arc(scale.x(entry[0].time), scale.y(entry[0].value), 4.5, 0, Math.PI * 2);
                ctx.fill(); ctx.stroke();
            });

            const lines = [{ text: formatClock(aHit ? aHit.time : (gHit ? gHit.time : t), true), color: '#0f172a', bold: true }];
            if (aHit) lines.push({ text: 'AMMA-2: ' + aHit.value.toFixed(1) + ' %', color: '#2563eb' });
            if (gHit) lines.push({ text: 'GRMS: ' + gHit.value.toFixed(1) + ' %', color: '#dc2626' });
            if (aHit && gHit) {
                const diff = aHit.value - gHit.value;
                lines.push({ text: 'Difference: ' + (diff >= 0 ? '+' : '') + diff.toFixed(1) + ' %', color: '#475569' });
            }
            if (lines.length === 1) lines.push({ text: 'No reading nearby', color: '#64748b' });

            ctx.font = '11px Segoe UI, Arial, sans-serif';
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            const boxW = Math.max(...lines.map(l => ctx.measureText(l.text).width)) + 20;
            const boxH = lines.length * 17 + 10;
            let bx = chartHoverX + 12;
            if (bx + boxW > width - 4) bx = chartHoverX - 12 - boxW;
            const by = plot.top + plot.height - boxH - 8;
            ctx.fillStyle = 'rgba(255,255,255,0.96)';
            ctx.strokeStyle = '#94a3b8';
            ctx.lineWidth = 1;
            ctx.fillRect(bx, by, boxW, boxH);
            ctx.strokeRect(bx + 0.5, by + 0.5, boxW, boxH);
            lines.forEach((line, i) => {
                ctx.fillStyle = line.color;
                ctx.font = (line.bold ? '600 ' : '') + '11px Segoe UI, Arial, sans-serif';
                ctx.fillText(line.text, bx + 10, by + 5 + i * 17 + 8.5);
            });
            ctx.restore();
        }

        function renderChart() {
            const canvasInfo = getCanvasContext();
            if (!canvasInfo || !chartState) return;
            const { ctx, width, height } = canvasInfo;
            const state = chartState;
            const { amma, grms, xMin, xMax, yMin, yMax } = state;

            const plot = { left: 60, right: 24, top: 20, bottom: 56 };
            plot.width = Math.max(1, width - plot.left - plot.right);
            plot.height = Math.max(1, height - plot.top - plot.bottom);
            const scale = {
                x: t => plot.left + ((t - xMin) / Math.max(1, xMax - xMin)) * plot.width,
                y: v => plot.top + (1 - ((v - yMin) / Math.max(1, yMax - yMin))) * plot.height
            };

            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, width, height);

            // Horizontal grid and Y-axis labels.
            ctx.font = '11px Segoe UI, Arial, sans-serif';
            ctx.textAlign = 'right';
            ctx.textBaseline = 'middle';
            for (let value = yMin; value <= yMax + 0.001; value += 20) {
                const y = scale.y(value);
                ctx.strokeStyle = '#e2e8f0'; ctx.lineWidth = 1;
                ctx.beginPath(); ctx.moveTo(plot.left, y); ctx.lineTo(plot.left + plot.width, y); ctx.stroke();
                ctx.fillStyle = '#475569'; ctx.fillText(value + '%', plot.left - 9, y);
            }

            // Vertical grid at round clock times (e.g. every 15 minutes), plus window edges.
            const step = pickTickStep(xMax - xMin);
            const dayStart = new Date(new Date(xMin).setHours(0, 0, 0, 0)).getTime();
            const firstTick = dayStart + Math.ceil((xMin - dayStart) / step) * step;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'top';
            for (let t = firstTick; t <= xMax + 1; t += step) {
                const x = scale.x(t);
                const major = ((t - dayStart) % (step * 2)) === 0 || step >= 30 * 60000;
                ctx.strokeStyle = major ? '#cbd5e1' : '#e2e8f0'; ctx.lineWidth = 1;
                ctx.beginPath(); ctx.moveTo(x, plot.top); ctx.lineTo(x, plot.top + plot.height); ctx.stroke();
                ctx.fillStyle = '#475569';
                ctx.fillText(formatClock(t, false), x, plot.top + plot.height + 8);
            }
            ctx.fillStyle = '#64748b';
            ctx.textAlign = 'left';
            ctx.fillText(formatClock(xMin, true), plot.left, plot.top + plot.height + 24);
            ctx.textAlign = 'right';
            ctx.fillText(formatClock(xMax, true), plot.left + plot.width, plot.top + plot.height + 24);

            // Reference thresholds; these are visual guides, not a declaration of compliance limits.
            [{ level: 100, color: '#64748b', label: '100%' }, { level: 125, color: '#d97706', label: '125% max' }].forEach(ref => {
                if (ref.level < yMin || ref.level > yMax) return;
                const y = scale.y(ref.level);
                ctx.save(); ctx.setLineDash([5, 4]);
                ctx.strokeStyle = ref.color; ctx.lineWidth = 1.2;
                ctx.beginPath(); ctx.moveTo(plot.left, y); ctx.lineTo(plot.left + plot.width, y); ctx.stroke();
                ctx.restore();
                ctx.fillStyle = ref.color;
                ctx.textAlign = 'right'; ctx.textBaseline = 'bottom';
                ctx.fillText(ref.label, plot.left + plot.width - 4, y - 2);
            });

            // Axis titles.
            ctx.save();
            ctx.translate(15, plot.top + plot.height / 2);
            ctx.rotate(-Math.PI / 2);
            ctx.fillStyle = '#334155'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.fillText('Positive Modulation (%)', 0, 0);
            ctx.restore();
            const dayLabel = new Date(xMin).toLocaleDateString([], { day: 'numeric', month: 'short', year: 'numeric' });
            ctx.fillStyle = '#334155'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
            ctx.fillText('Broadcast time (HH:MM) — ' + dayLabel, plot.left + plot.width / 2, height - 3);

            // Same axes for both datasets, making their timing and peaks directly comparable.
            drawSeries(ctx, grms, scale, plot, '#dc2626');
            drawSeries(ctx, amma, scale, plot, '#2563eb');

            drawLegend(ctx, plot, amma, grms);
            drawHover(ctx, plot, scale, state, width);
        }

        function renderSeriesSummary(amma, grms) {
            const box = document.getElementById('seriesSummary');
            if (!box) return;
            box.innerHTML = '';
            [
                { name: 'AMMA-2 (DB)', cls: 'amma-dot', stats: summarizeSeries(amma) },
                { name: 'GRMS', cls: 'grms-dot', stats: summarizeSeries(grms) }
            ].forEach(series => {
                const item = document.createElement('div');
                item.className = 'series-summary-item';
                const title = document.createElement('strong');
                const dot = document.createElement('i');
                dot.className = 'legend-dot ' + series.cls;
                title.appendChild(dot);
                title.appendChild(document.createTextNode(series.name));
                item.appendChild(title);
                const body = document.createElement('span');
                const s = series.stats;
                body.textContent = s
                    ? 'min ' + s.min.toFixed(1) + '%  ·  avg ' + s.avg.toFixed(1) + '%  ·  max ' + s.max.toFixed(1) +
                      '% at ' + formatClock(s.maxAt, true) + '  ·  ≥100%: ' + s.over100 + '  ·  ≥125%: ' + s.over125
                    : 'no data loaded';
                item.appendChild(body);
                box.appendChild(item);
            });
        }

        function drawModulationChart(ammaRows, grmsRows) {
            lastChartRows = { amma: ammaRows, grms: grmsRows };
            const message = document.getElementById('chartMessage');
            const toPoints = rows => rows.map(row => ({
                time: chartDate(row.timestamp)?.getTime(),
                value: Number(row.value)
            })).filter(p => Number.isFinite(p.time) && Number.isFinite(p.value))
              .sort((a, b) => a.time - b.time);
            const amma = toPoints(ammaRows);
            const grms = toPoints(grmsRows);
            renderSeriesSummary(amma, grms);

            if (!amma.length && !grms.length) {
                chartState = null;
                const canvasInfo = getCanvasContext();
                if (canvasInfo) {
                    canvasInfo.ctx.fillStyle = '#ffffff';
                    canvasInfo.ctx.fillRect(0, 0, canvasInfo.width, canvasInfo.height);
                }
                message.style.display = 'flex';
                message.innerText = 'Upload both files to compare positive modulation.';
                return;
            }
            message.style.display = 'none';

            const all = amma.concat(grms);
            const times = all.map(p => p.time);
            const win = windowBounds(Math.min(...times));
            // Fixed broadcast-window axis; widen only if a reading falls outside it.
            let xMin = Math.min(win.start, ...times);
            let xMax = Math.max(win.end, ...times);
            if (xMax <= xMin) { xMin -= 60000; xMax += 60000; }

            const values = all.map(p => p.value);
            // Keep a useful common scale and add headroom if either file exceeds 120%.
            const yMin = Math.min(0, Math.floor(Math.min(...values) / 20) * 20);
            let yMax = Math.max(140, Math.ceil(Math.max(...values) / 20) * 20);
            if (yMax <= yMin) yMax = yMin + 100;

            chartState = { amma, grms, xMin, xMax, yMin, yMax };
            renderChart();
        }

        async function refreshHistory() {
            const hours = 24;
            const message = document.getElementById('chartMessage');
            message.style.display = 'flex';
            message.innerText = 'Loading historical data...';

            try {
                const [historyResponse, grmsResponse, comparisonResponse] = await Promise.all([
                    fetch('/api/history?hours=' + encodeURIComponent(hours)),
                    fetch('/api/grms'),
                    fetch('/api/comparison?hours=' + encodeURIComponent(hours))
                ]);

                const history = await historyResponse.json();
                const grms = await grmsResponse.json();
                const comparison = await comparisonResponse.json();

                document.getElementById('ammaCount').innerText = history.length;
                document.getElementById('grmsCount').innerText = grms.length;
                document.getElementById('matchedCount').innerText = comparison.matched_points || 0;
                document.getElementById('maeValue').innerText =
                    comparison.mean_absolute_difference === null
                        ? '—'
                        : comparison.mean_absolute_difference.toFixed(2) + ' %';
                document.getElementById('correlationValue').innerText =
                    comparison.correlation === null
                        ? '—'
                        : comparison.correlation.toFixed(3);

                drawModulationChart(history, grms);

                const status = document.getElementById('uploadStatus');
                if (grms.length) {
                    status.innerText = 'GRMS data loaded: ' + grms.length + ' points.';
                } else {
                    status.innerText = 'No GRMS comparison file loaded.';
                }
            } catch (error) {
                console.error('Historical graph error:', error);
                message.style.display = 'flex';
                message.innerText = 'Unable to load historical data.';
            }
        }

        document.getElementById('ammaUploadForm').addEventListener('submit', async function(event) {
            event.preventDefault();
            const fileInput = document.getElementById('ammaFile');
            const status = document.getElementById('ammaUploadStatus');
            if (!fileInput.files.length) {
                status.innerText = 'Please choose an AMMA-2 CSV or XLSX file.';
                return;
            }
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            status.innerText = 'Uploading AMMA-2 data...';
            try {
                const response = await fetch('/upload-amma', { method: 'POST', body: formData });
                const result = await response.json();
                if (!response.ok) throw new Error(result.error || 'Upload failed.');
                status.innerText = 'AMMA-2 upload successful: ' + result.rows + ' points in 19:58–22:50.';
                await refreshHistory();
            } catch (error) {
                status.innerText = 'AMMA-2 upload failed: ' + error.message;
            }
        });

        document.getElementById('grmsUploadForm').addEventListener('submit', async function(event) {
            event.preventDefault();

            const fileInput = document.getElementById('grmsFile');
            const status = document.getElementById('uploadStatus');

            if (!fileInput.files.length) {
                status.innerText = 'Please choose a CSV or XLSX file.';
                return;
            }

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            status.innerText = 'Uploading GRMS data...';

            try {
                const response = await fetch('/upload-grms', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();

                if (!response.ok) {
                    throw new Error(result.error || 'Upload failed.');
                }

                status.innerText = 'GRMS upload successful: ' + result.rows + ' points.';
                await refreshHistory();
            } catch (error) {
                status.innerText = 'GRMS upload failed: ' + error.message;
            }
        });

        // Redraw from cached data on resize (no refetch), debounced so dragging stays smooth.
        let resizeTimer = null;
        window.addEventListener('resize', () => {
            clearTimeout(resizeTimer);
            resizeTimer = setTimeout(() => {
                drawModulationChart(lastChartRows.amma, lastChartRows.grms);
            }, 150);
        });

        function bindChartHover() {
            const canvas = document.getElementById('modulationChart');
            if (!canvas) return;
            const move = event => {
                chartHoverX = event.clientX - canvas.getBoundingClientRect().left;
                renderChart();
            };
            canvas.addEventListener('pointermove', move);
            canvas.addEventListener('pointerdown', move);
            canvas.addEventListener('pointerleave', () => {
                chartHoverX = null;
                renderChart();
            });
        }

        window.addEventListener('load', () => {
            structuralBuildLeds();
            evaluateInstrumentCompliance();
            startAutoRefresh();
            bindChartHover();
            refreshHistory();   // draw any already-uploaded data as soon as the page opens
        });
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    """Serves the main web interface dashboard."""
    return render_template_string(
        DASHBOARD_HTML,
        telemetry=live_telemetry,
        window_start=COMPARISON_START_TIME,
        window_end=COMPARISON_END_TIME,
    )


@app.route('/api/telemetry')
def api_telemetry():
    """Optional JSON API endpoint if you want to query data programmatically."""
    return jsonify(live_telemetry)



@app.route('/api/history')
def api_history():
    """Return AMMA-2 historical positive-modulation data."""
    try:
        hours = float(request.args.get("hours", "24"))
    except ValueError:
        hours = 24
    hours = max(1, min(hours, 24 * 30))
    return jsonify(read_comparison_amma())


@app.route('/api/grms')
def api_grms():
    """Return normalized GRMS positive-modulation comparison data."""
    return jsonify(_read_grms_rows())


@app.route('/api/comparison')
def api_comparison():
    """Return basic AMMA-vs-GRMS comparison statistics."""
    try:
        hours = float(request.args.get("hours", "24"))
    except ValueError:
        hours = 24
    hours = max(1, min(hours, 24 * 30))
    return jsonify(calculate_comparison(hours))


@app.route('/upload-amma', methods=['POST'])
def upload_amma():
    """Accept an AMMA-2 CSV/XLSX export containing Timestamp + DB Positive Modulation."""
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "No AMMA-2 file was selected."}), 400

    filename = uploaded.filename.lower()
    if not (filename.endswith(".csv") or filename.endswith(".xlsx")):
        return jsonify({"error": "Only CSV and XLSX files are supported."}), 400

    try:
        rows = _save_amma_upload(uploaded)
        return jsonify({"success": True, "rows": rows, "message": "AMMA-2 data normalized successfully."})
    except Exception as exc:
        logging.exception("AMMA-2 upload failed")
        return jsonify({"error": str(exc)}), 400


@app.route('/upload-grms', methods=['POST'])
def upload_grms():
    """Accept a GRMS CSV/XLSX export and normalize it for comparison."""
    uploaded = request.files.get("file")

    if not uploaded or not uploaded.filename:
        return jsonify({"error": "No file was selected."}), 400

    filename = uploaded.filename.lower()
    if not (filename.endswith(".csv") or filename.endswith(".xlsx")):
        return jsonify({"error": "Only CSV and XLSX files are supported."}), 400

    try:
        rows = _save_grms_upload(uploaded)
        return jsonify({
            "success": True,
            "rows": rows,
            "message": "GRMS data normalized successfully."
        })
    except Exception as exc:
        logging.exception("GRMS upload failed")
        return jsonify({"error": str(exc)}), 400


if __name__ == "__main__":
    # 1. Start the serial reader thread as a daemon (runs quietly in background)
    serial_thread = Thread(target=data_collection_worker, daemon=True)
    serial_thread.start()

    # 2. Run the Flask Web Application
    app.run(host='0.0.0.0', port=5000, debug=False)