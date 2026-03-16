"""Temporary script: print all unique Base Voltage (kVLL) values across CYME device reports."""
import csv
from pathlib import Path

REPORTS_DIR = Path("data/reports")

voltages = set()
for path in sorted(REPORTS_DIR.glob("CYME_Device_Report_*.csv")):
    with open(path, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            v = row.get("Base Voltage (kVLL)", "").strip()
            if v:
                voltages.add(v)

print("Unique Base Voltage (kVLL) values:")
for v in sorted(voltages, key=float):
    print(f"  {v}")
