"""
TWR Eswatini - Broadcast Schedule Checker
==========================================
Reads a broadcast schedule from a CSV file and checks
which broadcasts should currently be on air.

Author  : Learning project for TWR monitoring
Python  : 3.12+
Usage   : python schedule_checker.py
"""

import csv
from datetime import datetime


# ── CONFIG ──────────────────────────────────────────────────────────────────

SCHEDULE_FILE = "schedule.csv"   # CSV file in the same folder as this script


# ── FUNCTIONS ────────────────────────────────────────────────────────────────

def load_schedule(filename):
    """
    Read the CSV schedule file and return a list of broadcast entries.
    Each entry is a dictionary with keys matching the CSV column headers.
    """
    broadcasts = []

    with open(filename, newline="") as csvfile:
        reader = csv.DictReader(csvfile)       # DictReader turns each row into a dictionary
        for row in reader:
            broadcasts.append(row)

    return broadcasts


def parse_time(time_str):
    """
    Convert a time string like '06:30' into a Python time object
    so we can do proper time comparisons.
    """
    return datetime.strptime(time_str, "%H:%M").time()


def check_schedule(broadcasts, current_time):
    """
    Compare the current time against every broadcast in the schedule.
    Returns two lists:
      - on_air   : broadcasts that are currently active
      - upcoming : broadcasts starting in the next 30 minutes
    """
    on_air   = []
    upcoming = []

    for broadcast in broadcasts:
        start = parse_time(broadcast["start_time"])
        end   = parse_time(broadcast["end_time"])

        # Check if currently on air
        if start <= current_time <= end:
            on_air.append(broadcast)

        # Check if starting within the next 30 minutes
        else:
            # Calculate minutes until start
            start_dt   = datetime.combine(datetime.today(), start)
            current_dt = datetime.combine(datetime.today(), current_time)
            diff_minutes = (start_dt - current_dt).total_seconds() / 60

            if 0 < diff_minutes <= 30:
                upcoming.append((broadcast, int(diff_minutes)))

    return on_air, upcoming


def print_separator(char="─", width=55):
    """Print a visual divider line."""
    print(char * width)


def display_broadcast(broadcast, label=""):
    """Print one broadcast entry in a readable format."""
    print(f"  Frequency  : {broadcast['frequency_mhz']} MHz")
    print(f"  Time       : {broadcast['start_time']} – {broadcast['end_time']} UTC")
    print(f"  Language   : {broadcast['language']}")
    print(f"  Country    : {broadcast['target_country']}")
    print(f"  Transmitter: {broadcast['transmitter']}")
    if label:
        print(f"  Note       : {label}")


def run_check():
    """Main function — loads the schedule and prints the status report."""

    # Get the current time (hours and minutes only)
    now = datetime.now().time().replace(second=0, microsecond=0)

    print()
    print_separator("═")
    print("  TWR ESWATINI — BROADCAST SCHEDULE CHECKER")
    print_separator("═")
    print(f"  Check time : {now.strftime('%H:%M')} (local clock)")
    print(f"  Schedule   : {SCHEDULE_FILE}")
    print_separator()

    # Load schedule from CSV
    try:
        broadcasts = load_schedule(SCHEDULE_FILE)
    except FileNotFoundError:
        print(f"  ERROR: Could not find '{SCHEDULE_FILE}'")
        print("  Make sure the CSV file is in the same folder as this script.")
        return

    print(f"  Loaded {len(broadcasts)} broadcast entries from schedule.")
    print_separator()

    # Check what is on air and what is coming up
    on_air, upcoming = check_schedule(broadcasts, now)

    # ── ON AIR ──────────────────────────────────────────────────────────────
    if on_air:
        print(f"  ✔  ON AIR NOW ({len(on_air)} broadcast{'s' if len(on_air) > 1 else ''})")
        print_separator()
        for i, b in enumerate(on_air, 1):
            print(f"  [{i}]")
            display_broadcast(b)
            print()
    else:
        print("  ✘  NO BROADCASTS ON AIR RIGHT NOW")
        print_separator()

    # ── COMING UP ───────────────────────────────────────────────────────────
    if upcoming:
        print(f"  ⏱  COMING UP IN THE NEXT 30 MINUTES ({len(upcoming)} broadcast{'s' if len(upcoming) > 1 else ''})")
        print_separator()
        for i, (b, mins) in enumerate(upcoming, 1):
            print(f"  [{i}] Starts in {mins} minute{'s' if mins != 1 else ''}")
            display_broadcast(b)
            print()
    else:
        print("  No broadcasts starting in the next 30 minutes.")
        print_separator()

    # ── FULL SCHEDULE SUMMARY ────────────────────────────────────────────────
    print()
    print_separator("═")
    print("  FULL SCHEDULE SUMMARY")
    print_separator("═")
    print(f"  {'TIME':>13}  {'FREQ (MHz)':>10}  {'LANG':<12}  {'COUNTRY':<16}  TX")
    print_separator()
    for b in broadcasts:
        time_range = f"{b['start_time']}–{b['end_time']}"
        print(f"  {time_range:>13}  {b['frequency_mhz']:>10}  {b['language']:<12}  {b['target_country']:<16}  {b['transmitter']}")

    print_separator("═")
    print()


# ── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_check()
