"""Temporary script: print all unique Base Voltage (kVLL) values across CYME device reports
and write a CSV of each circuit's unique base voltages.
"""
import csv
from pathlib import Path

REPORTS_DIR = Path("data/reports")
OUTPUT_CSV  = Path("circuit_base_voltages.csv")

all_voltages: set[str] = set()
circuit_voltages: list[dict] = []

for path in sorted(REPORTS_DIR.glob("CYME_Device_Report_*.csv")):
    circuit = path.stem.replace("CYME_Device_Report_", "")
    voltages: set[str] = set()
    with open(path, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            v = row.get("Base Voltage (kVLL)", "").strip()
            if v:
                voltages.add(v)
                all_voltages.add(v)
    for v in sorted(voltages, key=float):
        circuit_voltages.append({"Circuit": circuit, "Base_Voltage_kVLL": v})

print("Unique Base Voltage (kVLL) values:")
for v in sorted(all_voltages, key=float):
    print(f"  {v}")

with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as fh:
    writer = csv.DictWriter(fh, fieldnames=["Circuit", "Base_Voltage_kVLL"])
    writer.writeheader()
    writer.writerows(circuit_voltages)

print(f"\nWrote {len(circuit_voltages)} rows to {OUTPUT_CSV}")
