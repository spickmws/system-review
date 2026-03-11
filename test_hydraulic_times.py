"""
test_hydraulic_times.py
=======================
Generates a table of slow-curve opening times for every hydraulic recloser
in protective_device_mapping.csv at two fault levels (1000 A and 2000 A).

Usage
-----
    python test_hydraulic_times.py
"""

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))
from protective_device_coordination import get_trip_time  # noqa: E402

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
_HERE = Path(__file__).parent
DEVICE_MAP_PATH = _HERE / "data" / "protective_device_mapping.csv"

# ---------------------------------------------------------------------------
# Same mapping used in main.py — equipment ID prefix → hydraulic curve type
# ---------------------------------------------------------------------------
HYDRAULIC_MAP: dict[str, str] = {
    "L":   "L",
    "V4L": "V4L_V4E",
    "V4H": "V4H",
    "4H":  "4H",
    "4E":  "4E",
    "E":   "E",
    "D":   "D",
    "CMR": "L",
}

FAULT_LEVELS = [1000, 2000]


def load_hydraulic_reclosers() -> list[dict]:
    """
    Return rows from protective_device_mapping.csv where Device Type is
    'Recloser' and Pickup is a plain numeric value (hydraulic).
    Adds 'prefix' and 'curve_type' derived fields.
    """
    reclosers = []
    with open(DEVICE_MAP_PATH, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            if row["Device Type"].strip() != "Recloser":
                continue
            pickup_str = row["Pickup"].strip()
            try:
                pickup = float(pickup_str)
            except ValueError:
                continue  # skip Electronic / TS / blank
            equip_id = row["Equipment ID"].strip()
            prefix = equip_id.split("_")[0]
            curve_type = HYDRAULIC_MAP.get(prefix)
            if curve_type is None:
                curve_type = "UNKNOWN"
            reclosers.append({
                "Equipment ID": equip_id,
                "Pickup_A":     pickup,
                "Prefix":       prefix,
                "Curve Type":   curve_type,
            })
    return sorted(reclosers, key=lambda r: (r["Prefix"], r["Pickup_A"]))


def fmt_time(curve_type: str, pickup_a: float, fault_a: float) -> str:
    """Return formatted trip time string, or an error note.

    The hydraulic curves are indexed by multiple-of-pickup (fault_a / pickup),
    matching the same convention used in main.py.
    """
    if curve_type == "UNKNOWN":
        return "no curve map"
    if pickup_a <= 0:
        return "zero pickup"
    multiple = fault_a / pickup_a
    try:
        t = get_trip_time(device="hydraulic", curve_type=curve_type,
                          i=multiple, curve="slow")
        return f"{t:.4f}"
    except ValueError as exc:
        msg = str(exc)
        if "outside the data range" in msg:
            return "out of range"
        return f"err: {msg[:30]}"


def main() -> None:
    reclosers = load_hydraulic_reclosers()

    # Column widths
    col_id    = max(len("Equipment ID"),   max(len(r["Equipment ID"])  for r in reclosers))
    col_pu    = max(len("Pickup (A)"),     6)
    col_curve = max(len("Curve Type"),     max(len(r["Curve Type"])    for r in reclosers))
    col_t     = 12  # wide enough for time values

    fault_headers = [f"{fa} A" for fa in FAULT_LEVELS]

    # Header
    header = (
        f"{'Equipment ID':<{col_id}}  "
        f"{'Pickup (A)':>{col_pu}}  "
        f"{'Curve Type':<{col_curve}}"
        + "".join(f"  {h:>{col_t}}" for h in fault_headers)
    )
    separator = "-" * len(header)

    print()
    print("Hydraulic Recloser Slow-Curve Opening Times")
    print(separator)
    print(header)
    print(separator)

    for r in reclosers:
        times = [fmt_time(r["Curve Type"], r["Pickup_A"], fa)
                 for fa in FAULT_LEVELS]
        row = (
            f"{r['Equipment ID']:<{col_id}}  "
            f"{r['Pickup_A']:>{col_pu}.0f}  "
            f"{r['Curve Type']:<{col_curve}}"
            + "".join(f"  {t:>{col_t}}" for t in times)
        )
        print(row)

    print(separator)
    print(f"\n{len(reclosers)} recloser(s) | fault levels: "
          f"{', '.join(str(fa) + ' A' for fa in FAULT_LEVELS)} | curve: slow\n")


if __name__ == "__main__":
    main()
