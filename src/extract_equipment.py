"""
extract_equipment.py
--------------------
Scans a folder for all files matching 'Feeder_Overview*.csv', reads each file,
and produces a single deduplicated CSV containing every unique Equipment Number
and its corresponding Device Type.

Usage:
    python extract_equipment.py --input-dir /path/to/folder --output unique_equipment.csv

Arguments:
    --input-dir   Directory containing the Feeder_Overview CSV files (required)
    --output      Output CSV file path (default: unique_equipment.csv)
    --encoding    File encoding to use when reading CSVs (default: utf-8-sig)
    --verbose     Print per-file progress to stdout
"""

import argparse
import csv
import glob
import logging
import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
FILE_GLOB_PATTERN = "CYME*.csv"
DEVICE_TYPE_COL   = "Device Type"
EQUIPMENT_ID_COL  = "Equipment ID"
DEFAULT_OUTPUT     = "unique_equipment.csv"
DEFAULT_ENCODING   = "utf-8-sig"   # handles BOM that Excel sometimes adds


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        format="%(levelname)s: %(message)s",
        level=level,
    )


def find_input_files(input_dir: str) -> list[Path]:
    """Return all Feeder_Overview CSV files in *input_dir*, sorted by name."""
    pattern = os.path.join(input_dir, FILE_GLOB_PATTERN)
    files = sorted(Path(p) for p in glob.glob(pattern))
    return files


def extract_from_file(
    filepath: Path,
    encoding: str,
) -> tuple[list[tuple[str, str]], list[str]]:
    """
    Parse one CSV file and return:
      - list of (equipment_id, device_type) tuples found in the file
      - list of warning strings for any rows that were skipped
    """
    records: list[tuple[str, str]] = []
    warnings: list[str] = []

    try:
        with filepath.open(newline="", encoding=encoding, errors="replace") as fh:
            reader = csv.DictReader(fh)

            # Validate required columns exist
            if reader.fieldnames is None:
                warnings.append(f"{filepath.name}: file appears to be empty — skipped")
                return records, warnings

            missing = {DEVICE_TYPE_COL, EQUIPMENT_ID_COL} - set(reader.fieldnames)
            if missing:
                warnings.append(
                    f"{filepath.name}: missing required column(s) {missing} — skipped"
                )
                return records, warnings

            for line_num, row in enumerate(reader, start=2):  # 2 because row 1 is header
                eq_id  = row[EQUIPMENT_ID_COL].strip()
                dev_type = row[DEVICE_TYPE_COL].strip()

                if not eq_id:
                    warnings.append(
                        f"{filepath.name} line {line_num}: empty Equipment Number — row skipped"
                    )
                    continue

                records.append((eq_id, dev_type))

    except (OSError, UnicodeDecodeError) as exc:
        warnings.append(f"{filepath.name}: could not read file ({exc}) — skipped")

    return records, warnings


def build_unique_equipment(
    files: list[Path],
    encoding: str,
    verbose: bool,
) -> dict[str, str]:
    """
    Process every file and return a dict mapping equipment_id → device_type.
    If the same ID appears with different device types across files, the first
    occurrence wins and a warning is logged.
    """
    equipment: dict[str, str] = {}   # equipment_id → device_type
    conflict_ids: set[str] = set()

    for filepath in files:
        records, warnings = extract_from_file(filepath, encoding)

        for msg in warnings:
            logging.warning(msg)

        new_count = 0
        for eq_id, dev_type in records:
            if eq_id not in equipment:
                equipment[eq_id] = dev_type
                new_count += 1
            else:
                existing = equipment[eq_id]
                if existing != dev_type and eq_id not in conflict_ids:
                    logging.warning(
                        "Equipment ID '%s' has conflicting device types: "
                        "'%s' (kept) vs '%s' (in %s, ignored).",
                        eq_id, existing, dev_type, filepath.name,
                    )
                    conflict_ids.add(eq_id)

        logging.debug("  %s → %d records (%d new unique IDs)", filepath.name, len(records), new_count)

    return equipment


def write_output(equipment: dict[str, str], output_path: str) -> None:
    """Write the deduplicated equipment dict to a CSV file, sorted by Equipment Number."""
    rows = sorted(equipment.items(), key=lambda x: x[0])   # sort by Equipment Number

    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow([EQUIPMENT_ID_COL, DEVICE_TYPE_COL])
        writer.writerows(rows)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract unique equipment IDs and device types from all "
            "Feeder_Overview CSV files in a directory."
        )
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        metavar="DIR",
        help="Directory containing Feeder_Overview*.csv files",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        metavar="FILE",
        help=f"Output CSV path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--encoding",
        default=DEFAULT_ENCODING,
        help=f"Input file encoding (default: {DEFAULT_ENCODING})",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print per-file debug information",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    configure_logging(args.verbose)

    # Validate input directory
    if not os.path.isdir(args.input_dir):
        logging.error("Input directory not found: %s", args.input_dir)
        return 1

    # Discover files
    files = find_input_files(args.input_dir)
    if not files:
        logging.error(
            "No files matching '%s' found in: %s",
            FILE_GLOB_PATTERN,
            args.input_dir,
        )
        return 1

    logging.info("Found %d file(s) to process in '%s'.", len(files), args.input_dir)

    # Process
    equipment = build_unique_equipment(files, args.encoding, args.verbose)

    if not equipment:
        logging.error("No equipment records were extracted. Output not written.")
        return 1

    # Write results
    write_output(equipment, args.output)
    logging.info(
        "Done. %d unique equipment IDs written to '%s'.",
        len(equipment),
        args.output,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
