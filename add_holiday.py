#!/usr/bin/env python3
"""Add, list, or remove holiday/leave entries in config.yaml.

No YAML editing required.

Usage:
    python add_holiday.py 2026-10-02 "Office Holiday"   # add a holiday
    python add_holiday.py --list                          # show all holidays
    python add_holiday.py --remove 2026-10-02             # remove a holiday

The date must fall inside one of the configured timesheet ranges, otherwise
the generator will ignore it. Run `python add_holiday.py --list` to see the
active ranges.
"""

import argparse
import datetime
import sys
from pathlib import Path

import yaml

CONFIG_PATH = Path(__file__).parent / "config.yaml"


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"Config not found: {CONFIG_PATH}")
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def save_config(cfg: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, sort_keys=False, allow_unicode=True, default_flow_style=False)


def parse_date(value: str) -> datetime.date:
    try:
        return datetime.date.fromisoformat(value)
    except ValueError:
        raise ValueError(f"Invalid date '{value}' — use YYYY-MM-DD")


def in_range(date: datetime.date, cfg: dict) -> bool:
    for ts in cfg.get("timesheets", []):
        start = datetime.date.fromisoformat(ts["start"])
        end = datetime.date.fromisoformat(ts["end"])
        if start <= date <= end:
            return True
    return False


def cmd_add(args, cfg: dict) -> int:
    date = parse_date(args.date)
    if not in_range(date, cfg):
        ranges = ", ".join(
            f"{ts['start']} to {ts['end']}" for ts in cfg.get("timesheets", [])
        )
        print(
            f"WARNING: {date} is not inside any timesheet range ({ranges}). "
            "The generator will ignore it.",
            file=sys.stderr,
        )
    leaves = cfg.setdefault("leaves", [])
    for entry in leaves:
        if entry["date"] == date.isoformat():
            print(f"Holiday on {date} already exists: {entry['type']!r}")
            return 0
    leaves.append({"date": date.isoformat(), "type": args.type})
    save_config(cfg)
    print(f"Added: {date} -> {args.type!r}")
    return 0


def cmd_list(args, cfg: dict) -> int:
    leaves = cfg.get("leaves", [])
    if not leaves:
        print("No holidays configured.")
        return 0
    print(f"{len(leaves)} holiday(s):")
    for entry in leaves:
        print(f"  {entry['date']}  {entry['type']}")
    return 0


def cmd_remove(args, cfg: dict) -> int:
    target = parse_date(args.date).isoformat()
    leaves = cfg.get("leaves", [])
    before = len(leaves)
    cfg["leaves"] = [e for e in leaves if e["date"] != target]
    if len(cfg["leaves"]) == before:
        print(f"No holiday on {target} to remove.")
        return 0
    save_config(cfg)
    print(f"Removed holiday on {target}.")
    return 0


def cmd_set_range(args, cfg: dict) -> int:
    """Replace the two timesheet ranges with new ones."""
    new_ranges = [
        {"name": "Timesheet_1", "start": args.start1, "end": args.end1},
        {"name": "Timesheet_2", "start": args.start2, "end": args.end2},
    ]
    cfg["timesheets"] = new_ranges
    save_config(cfg)
    print("Timesheet ranges updated:")
    for ts in new_ranges:
        print(f"  {ts['name']}: {ts['start']} -> {ts['end']}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Add, list, or remove holidays in config.yaml without editing YAML by hand."
    )
    parser.add_argument("date", nargs="?", help="YYYY-MM-DD")
    parser.add_argument("type", nargs="?", help="Leave/holiday type, e.g. 'Office Holiday'")
    parser.add_argument("--list", action="store_true", help="List all configured holidays")
    parser.add_argument("--remove", metavar="DATE", help="Remove the holiday on DATE")
    parser.add_argument(
        "--set-range",
        nargs=4,
        metavar=("START1", "END1", "START2", "END2"),
        help="Replace both timesheet ranges, e.g. --set-range 2026-10-01 2026-10-31 2026-09-26 2026-10-26",
    )
    args = parser.parse_args()

    cfg = load_config()

    if args.set_range:
        return cmd_set_range(argparse.Namespace(
            start1=args.set_range[0], end1=args.set_range[1],
            start2=args.set_range[2], end2=args.set_range[3],
        ), cfg)
    if args.list:
        return cmd_list(args, cfg)
    if args.remove:
        return cmd_remove(argparse.Namespace(date=args.remove), cfg)
    if args.date and args.type:
        return cmd_add(args, cfg)

    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())