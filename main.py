"""
main.py
=======
DEC Protection Coordination Review

Processes CYME_Device_Report and relay/recloser settings files to evaluate
whether each upstream/downstream protective device pair meets minimum
coordination margins (breaker, recloser, fuse zones).

Writes a single consolidated violations_report.csv.

Usage
-----
    python main.py

Configure DATA_DIR / RECLOSER_DB_PATH below, or override via environment:
    COORD_REPORTS_DIR   path to folder containing CYME_Device_Report_* and
                        Feeder_Relay_Settings_* files
    COORD_RECLOSER_DB   path to RecloserDatabase.xlsx
"""

import csv
import os
import re
import sys
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Locate the protective_device_coordination module
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent / "src"))
from protective_device_coordination import get_trip_time  # noqa: E402

# ---------------------------------------------------------------------------
# Paths  (override via environment variables if needed)
# ---------------------------------------------------------------------------
_HERE = Path(__file__).parent
DATA_DIR = _HERE / "data"
REPORTS_DIR = Path(os.environ.get("COORD_REPORTS_DIR", str(DATA_DIR / "reports")))
RECLOSER_DB_PATH = Path(
    os.environ.get("COORD_RECLOSER_DB", str(DATA_DIR / "references/RecloserDatabase.xlsx"))
)
DEVICE_MAP_PATH    = DATA_DIR / "protective_device_mapping.csv"
TRIPSAVER_DB_PATH  = DATA_DIR / "references" / "tripsavers.csv"
OUTPUT_PATH = _HERE / "violations_report.csv"
DISCREPANCY_PATH = _HERE / "voltage_discrepancies.csv"
INDETERMINATE_PATH = _HERE / "indeterminate_pairs.csv"

_TS_PART_MAP: dict[str, str] = {"80T": "TS80T", "100T": "TS100T", "150E": "TS150E"}

_VOLTAGE_CLASS_MAP: dict[str, set[float]] = {
    "12": {12.47, 12.5},
    "04": {2.4, 4.16},
    "24": {23.9, 24.94},
    "34": {34.5},
}

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Supported CYME Device Types that represent protective devices
_PROTECTIVE_TYPES = {"Breaker", "Fuse", "Recloser"}

# Minimum coordination margins (seconds).
# Key: (upstream_class, downstream_class)
THRESHOLDS: dict[tuple[str, str], float] = {
    ("breaker",             "electronic_recloser"): 0.20,
    ("breaker",             "hydraulic_recloser"):  0.20,
    ("breaker",             "tripsaver"):            0.20,
    ("breaker",             "fuse"):                 0.00,
    ("electronic_recloser", "electronic_recloser"): 0.22,
    ("electronic_recloser", "hydraulic_recloser"):  0.22,
    ("electronic_recloser", "tripsaver"):            0.22,
    ("electronic_recloser", "fuse"):                 0.00,
    ("hydraulic_recloser",  "electronic_recloser"): 0.22,
    ("hydraulic_recloser",  "hydraulic_recloser"):  0.22,
    ("hydraulic_recloser",  "tripsaver"):            0.22,
    ("hydraulic_recloser",  "fuse"):                 0.00,
    ("tripsaver",           "electronic_recloser"): 0.22,
    ("tripsaver",           "hydraulic_recloser"):  0.22,
    ("tripsaver",           "tripsaver"):            0.22,
    ("tripsaver",           "fuse"):                 0.00,
    ("fuse",                "fuse"):                 0.00,
}

# Groups of electronic reclosers excluded from coordination checks
IGNORED_RECLOSER_GROUPS: set[int] = {5, 7, 8}

# Column name in RecloserDatabase.xlsx that holds the group number
RECLOSER_GROUP_COLUMN = "Group"

# Physical interrupter opening time added to all electronic recloser trip times (seconds)
ELECTRONIC_RECLOSER_OPENING_TIME_S: float = 0.064

# Hydraulic recloser Equipment ID prefix → curve type in hydraulic_i_t.csv
HYDRAULIC_MAP: dict[str, str] = {
    "L":   "L",
    "V4L": "V4L_V4E",
    "V4H": "V4H",
    "4H":  "4H",
    "4E":  "4E",
    "E":   "E",
    "D":   "D",
    "CMR": "L",   # CMR uses same curve as L-type
}

# Human-readable labels for the Check_Type column
_CLASS_LABEL: dict[str, str] = {
    "breaker":             "Breaker",
    "electronic_recloser": "Recloser",
    "hydraulic_recloser":  "Recloser",
    "tripsaver":           "TripSaver",
    "fuse":                "Fuse",
    "unknown":             "Unknown",
}

# Output column order
OUTPUT_FIELDNAMES = [
    "Circuit",
    "Upstream_Equipment_Number",
    "Upstream_Equipment_ID",
    "Upstream_Device_Type",
    "Downstream_Equipment_Number",
    "Downstream_Equipment_ID",
    "Downstream_Device_Type",
    "Check_Type",
    "Phase_Fault_Current_A",
    "Phase_Upstream_Time_s",
    "Phase_Downstream_Time_s",
    "Phase_Margin_s",
    "Phase_Required_Margin_s",
    "Phase_Violation",
    "Ground_Fault_Current_A",
    "Ground_Upstream_Time_s",
    "Ground_Downstream_Time_s",
    "Ground_Margin_s",
    "Ground_Required_Margin_s",
    "Ground_Violation",
    "Upstream_Phase_Pickup_A",
    "Upstream_Phase_Curve",
    "Upstream_Phase_TD",
    "Upstream_Ground_Pickup_A",
    "Upstream_Ground_Curve",
    "Upstream_Ground_TD",
    "Downstream_Phase_Pickup_A",
    "Downstream_Phase_Curve",
    "Downstream_Phase_TD",
    "Downstream_Ground_Pickup_A",
    "Downstream_Ground_Curve",
    "Downstream_Ground_TD",
    "Upstream_Total_Customers",
    "Downstream_Total_Customers",
    "Customers_Erroneously_Disconnected",
    "Upstream_Device_Downstream_A_PH_cKVA",
    "Upstream_Device_Downstream_B_PH_cKVA",
    "Upstream_Device_Downstream_C_PH_cKVA",
    "Upstream_Device_Total_Downstream_3PH_cKVA",
    "Downstream_Device_Downstream_A_PH_cKVA",
    "Downstream_Device_Downstream_B_PH_cKVA",
    "Downstream_Device_Downstream_C_PH_cKVA",
    "Downstream_Device_Total_Downstream_3PH_cKVA",
    "Notes",
]

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_device_map() -> dict[str, dict]:
    """
    Load protective_device_mapping.csv.
    Returns {equipment_id: {'Device Type': str, 'Pickup': str}}.
    """
    result: dict[str, dict] = {}
    with open(DEVICE_MAP_PATH, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            eid = row["Equipment ID"].strip()
            result[eid] = {
                "Device Type": row["Device Type"].strip(),
                "Pickup":      row["Pickup"].strip(),
            }
    return result


def load_tripsaver_db() -> dict[str, str]:
    """
    Load data/references/tripsavers.csv.
    Returns {rec_id_str: curve_type} e.g. {'581519278': 'TS80T'}.
    Curve type is parsed from the unit_cu part number string.
    """
    result: dict[str, str] = {}
    with open(TRIPSAVER_DB_PATH, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            rec_id  = row["rec_id"].strip()
            unit_cu = row["unit_cu"].strip()
            for token in unit_cu.split("-"):
                curve = _TS_PART_MAP.get(token.upper())
                if curve:
                    result[rec_id] = curve
                    break
    return result


def load_recloser_db() -> dict[int, dict]:
    """
    Read the 'Reclosers' sheet of RecloserDatabase.xlsx.
    Returns {int(FID): {column: value, ...}}.
    """
    wb = openpyxl.load_workbook(RECLOSER_DB_PATH, read_only=True, data_only=True)
    ws = wb["Reclosers"]
    row_iter = ws.iter_rows(values_only=True)
    header = [str(h) for h in next(row_iter)]
    result: dict[int, dict] = {}
    for row in row_iter:
        if row[0] is None:
            continue
        rec = dict(zip(header, row))
        try:
            fid = int(rec["FID"])
        except (ValueError, TypeError):
            continue
        result[fid] = rec
    wb.close()
    return result


def _parse_relay_settings(filepath: Path) -> dict[str, str]:
    """
    Parse a Feeder_Relay_Settings CSV (key-value row format).
    Returns {setting_name: value_string}.
    """
    settings: dict[str, str] = {}
    with open(filepath, encoding="utf-8-sig") as fh:
        for row in csv.reader(fh):
            if len(row) >= 2 and row[0].strip():
                settings[row[0].strip()] = row[1].strip()
    return settings


def load_relay_settings() -> dict[str, dict]:
    """
    Scan all Feeder_Relay_Settings_########.csv files.
    Returns {circuit_number: settings_dict}.
    """
    result: dict[str, dict] = {}
    for path in REPORTS_DIR.glob("Feeder_Relay_Settings_*.csv"):
        circuit = path.stem.rsplit("_", 1)[-1]
        result[circuit] = _parse_relay_settings(path)
    return result


def load_cyme_report(path: Path) -> dict[str, dict]:
    """
    Read a CYME_Device_Report CSV.
    Returns {equipment_number: row_dict} for Breaker/Fuse/Recloser rows
    that have Status == 'ConnectedClose'.
    """
    devices: dict[str, dict] = {}
    with open(path, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            if row.get("Device Type", "").strip() not in _PROTECTIVE_TYPES:
                continue
            if row.get("Status", "").strip() != "ConnectedClose":
                continue
            equip_num = row.get("Equipment Number", "").strip()
            if equip_num:
                devices[equip_num] = row
    return devices


def _check_voltage_class(circuit: str, devices: dict) -> list[dict]:
    """
    Return discrepancy rows if any device base voltage doesn't match the
    circuit's voltage class (positions [4:6] of the 8-char circuit ID).
    """
    voltage_class = circuit[4:6] if len(circuit) == 8 else None
    expected = _VOLTAGE_CLASS_MAP.get(voltage_class) if voltage_class else None
    if expected is None:
        return []

    mismatched: set[float] = set()
    for row in devices.values():
        bv = _safe_float(row.get("Base Voltage (kVLL)"), None)
        if bv is None or bv not in expected:
            mismatched.add(bv if bv is not None else 0.0)

    return [{"Circuit": circuit, "Base_Voltage_kVLL": v} for v in sorted(mismatched)]


# ---------------------------------------------------------------------------
# Device classification helpers
# ---------------------------------------------------------------------------

def extract_fid(equipment_number: str) -> str:
    """Extract FID from an equipment number (last token after the final '_')."""
    return equipment_number.rsplit("_", 1)[-1] if "_" in equipment_number else equipment_number


def _recloser_group(equipment_number: str, recloser_db: dict) -> int | None:
    """Return the group number for an electronic recloser, or None if not found."""
    fid_str = extract_fid(equipment_number)
    try:
        fid = int(fid_str)
    except ValueError:
        return None
    rec = recloser_db.get(fid)
    if rec is None:
        return None
    try:
        return int(rec[RECLOSER_GROUP_COLUMN])
    except (KeyError, TypeError, ValueError):
        return None


def _is_ignored_recloser(equipment_number: str, equipment_id: str,
                          cyme_type: str, device_map: dict,
                          recloser_db: dict) -> bool:
    """True if this device is an electronic recloser in an ignored group."""
    if classify_device(equipment_id, cyme_type, device_map) != "electronic_recloser":
        return False
    grp = _recloser_group(equipment_number, recloser_db)
    return grp in IGNORED_RECLOSER_GROUPS


def _effective_upstream(equipment_number: str, devices: dict,
                         device_map: dict, recloser_db: dict) -> dict | None:
    """
    Follow the upstream chain from equipment_number, skipping any electronic
    reclosers whose group is in IGNORED_RECLOSER_GROUPS.
    Returns the first non-ignored upstream device row, or None.
    """
    visited: set[str] = set()
    current_num = equipment_number
    while True:
        row = devices.get(current_num)
        if row is None:
            return None
        upstream_key = row.get("Upstream Protective Device", "").strip()
        if not upstream_key:
            return None
        if upstream_key in visited:
            return None  # cycle guard
        visited.add(upstream_key)
        u_row = devices.get(upstream_key)
        if u_row is None:
            return None
        u_equip_num = u_row.get("Equipment Number", "").strip()
        u_equip_id  = u_row.get("Equipment ID", "").strip()
        u_cyme_type = u_row.get("Device Type", "").strip()
        if _is_ignored_recloser(u_equip_num, u_equip_id, u_cyme_type,
                                 device_map, recloser_db):
            current_num = upstream_key  # skip and keep walking up
        else:
            return u_row               # found a non-ignored upstream


def classify_device(equipment_id: str, device_type_from_cyme: str,
                    device_map: dict) -> str:
    """
    Return one of:
        'breaker', 'electronic_recloser', 'hydraulic_recloser',
        'tripsaver', 'fuse', 'unknown'
    """
    # CYME 'Device Type' column is the authoritative source for Breaker
    if device_type_from_cyme == "Breaker":
        return "breaker"

    entry = device_map.get(equipment_id)
    if entry is None:
        return "unknown"

    dev_type = entry["Device Type"]
    pickup   = entry["Pickup"]

    if dev_type == "Breaker":
        return "breaker"
    elif dev_type == "Recloser":
        if pickup.upper().startswith("TS"):
            return "tripsaver"
        elif pickup == "Electronic":
            return "electronic_recloser"
        else:
            try:
                float(pickup)
                return "hydraulic_recloser"
            except ValueError:
                return "unknown"
    elif dev_type == "Fuse":
        return "fuse"

    return "unknown"


def get_fuse_curve_type(equipment_id: str, base_kv: float) -> str | None:
    """
    Map a fuse Equipment ID to a curve type key in fuses_i_t.csv.
    Returns None if the type cannot be resolved.

    Examples
    --------
    T_65A              → 65T
    K_20A              → 20K
    E_50A              → 50E_14.4kV   (base_kv <= 15)
    SMU-20_(E)_50A     → 50E_14.4kV
    SM-4_(E)_65A       → 65E_14.4kV
    """
    voltage_class = "14.4kV" if base_kv <= 15.0 else "25kV"

    # Rating is always the last token, e.g. '65A' → '65'
    last = equipment_id.rsplit("_", 1)[-1]
    if last.upper().endswith("A") and last[:-1].isdigit():
        rating = last[:-1]
    else:
        return None

    # E-type: standalone 'E_*', or embedded '_(E)_*' (SMU-20, SM-4, SMD-20)
    if "_(E)_" in equipment_id or equipment_id.upper().startswith("E_"):
        return f"{rating}E_{voltage_class}"
    # T-type
    if equipment_id.upper().startswith("T_") or equipment_id.upper().startswith("STANDARD_"):
        return f"{rating}T"
    # K-type
    if equipment_id.upper().startswith("K_"):
        return f"{rating}K"

    return None  # D_, ES_, UNK_, N_ → unresolvable


# ---------------------------------------------------------------------------
# Trip-time computation helpers
# ---------------------------------------------------------------------------

def _safe_float(value, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


_U_CURVE_PATTERN = re.compile(r'^U[1-5]$', re.IGNORECASE)

_RELAY_KEY_ALIASES: dict[str, list[str]] = {
    "51P1P":  ["51P1P", "51PP"],
    "51P1C":  ["51P1C", "51PC"],
    "51P1TD": ["51P1TD", "51PTD"],
    "51G1P":  ["51G1P", "51NP"],
    "51G1C":  ["51G1C", "51NC"],
    "51G1TD": ["51G1TD", "51NTD"],
}


def _relay_get(s: dict, key: str) -> str:
    """Return the first matching alias value from a relay settings dict."""
    for alias in _RELAY_KEY_ALIASES[key]:
        if alias in s:
            return s[alias]
    raise KeyError(key)


def _breaker_time(circuit: str, fault_a: float, element: str,
                  relay_settings: dict) -> tuple[float | None, str, dict]:
    """
    Compute breaker trip time from Feeder_Relay_Settings.
    element: 'phase' or 'ground'
    Returns (seconds, note, settings).
    """
    s = relay_settings.get(circuit)
    if s is None:
        return None, f"No relay settings for circuit {circuit}", {}

    try:
        ctr = float(s["CTR"])
        if element == "phase":
            pickup = float(_relay_get(s, "51P1P")) * ctr
            curve  = _relay_get(s, "51P1C").strip()
            td     = float(_relay_get(s, "51P1TD"))
        else:
            pickup = float(_relay_get(s, "51G1P")) * ctr
            curve  = _relay_get(s, "51G1C").strip()
            td     = float(_relay_get(s, "51G1TD"))
    except (KeyError, ValueError) as exc:
        return None, f"Relay setting error ({exc})", {}

    if pickup <= 0:
        return None, "Breaker pickup is zero", {}

    multiple = fault_a / pickup
    try:
        if _U_CURVE_PATTERN.match(curve):
            t = get_trip_time(device="ucurve", curve_type=curve, i=multiple, time_dial=td)
        else:
            t = get_trip_time(device="curve", curve_type=curve, i=multiple) * td
        return t, "", {"pickup": pickup, "curve": curve, "td": td}
    except ValueError as exc:
        return None, str(exc), {}


def _electronic_recloser_time(equipment_number: str, fault_a: float, element: str,
                               recloser_db: dict) -> tuple[float | None, str, dict]:
    """
    Compute electronic recloser slow-curve trip time from RecloserDatabase.
    Returns (seconds, note, settings).
    """
    fid_str = extract_fid(equipment_number)
    try:
        fid = int(fid_str)
    except ValueError:
        return None, f"Cannot parse FID from '{equipment_number}'", {}

    rec = recloser_db.get(fid)
    if rec is None:
        return None, f"FID {fid} not in RecloserDatabase", {}

    try:
        ctr = float(rec["CTR"])
        if element == "phase":
            pickup = float(rec["ph"]) * ctr
            curve  = rec["phSC"]
            td     = float(rec["phSTD"])
        else:
            pickup = float(rec["gPU"]) * ctr
            curve  = rec["gSC"]
            td     = float(rec["gSTD"])
    except (KeyError, TypeError, ValueError) as exc:
        return None, f"RecloserDB field error ({exc})", {}

    if pickup <= 0:
        return None, "Recloser pickup is zero", {}

    multiple = fault_a / pickup
    try:
        if isinstance(curve, str):
            # U-curve (e.g. 'U4')
            t = get_trip_time(device="ucurve", curve_type=curve.strip(),
                              i=multiple, time_dial=td)
        else:
            # Numeric SEL curve code — multiply result by time dial
            t = get_trip_time(device="curve", curve_type=str(int(curve)),
                              i=multiple) * td
        return t + ELECTRONIC_RECLOSER_OPENING_TIME_S, "", {"pickup": pickup, "curve": str(curve).strip(), "td": td}
    except ValueError as exc:
        return None, str(exc), {}


def _hydraulic_recloser_time(equipment_id: str, fault_a: float,
                              device_map: dict) -> tuple[float | None, str, dict]:
    """
    Compute hydraulic recloser slow-curve trip time.
    Returns (seconds, note, settings).
    """
    entry = device_map.get(equipment_id)
    if entry is None:
        return None, f"'{equipment_id}' not in device map", {}

    try:
        pickup = float(entry["Pickup"])
    except ValueError:
        return None, f"Non-numeric pickup for '{equipment_id}'", {}

    prefix = equipment_id.split("_")[0]
    curve_type = HYDRAULIC_MAP.get(prefix)
    if curve_type is None:
        return None, f"No hydraulic curve mapping for prefix '{prefix}'", {}

    if pickup <= 0:
        return None, "Hydraulic pickup is zero", {}

    multiple = fault_a / pickup
    try:
        t = get_trip_time(device="hydraulic", curve_type=curve_type,
                          i=multiple, curve="slow")
        return t, "", {"pickup": pickup, "curve": curve_type, "td": None}
    except ValueError as exc:
        return None, str(exc), {}


def _tripsaver_time(fault_a: float, curve_type: str) -> tuple[float | None, str, dict]:
    """
    Compute TripSaver trip time using the curve specified in the device mapping.
    i is in primary amps.
    Returns (seconds, note, settings).
    """
    try:
        t = get_trip_time(device="ts", curve_type=curve_type, i=fault_a)
        return t, "", {"pickup": None, "curve": curve_type, "td": None}
    except ValueError as exc:
        return None, str(exc), {}


def _fuse_time(equipment_id: str, fault_a: float, base_kv: float,
               fuse_curve: str) -> tuple[float | None, str, dict]:
    """
    Compute fuse trip time.
    fuse_curve: 'Melting' (upstream role) or 'Clearing' (downstream role).
    Returns (seconds, note, settings).
    """
    curve_type = get_fuse_curve_type(equipment_id, base_kv)
    if curve_type is None:
        return None, f"Cannot determine fuse curve type for '{equipment_id}'", {}
    try:
        t = get_trip_time(device="fuse", curve_type=curve_type,
                          i=fault_a, curve=fuse_curve)
        return t, "", {"pickup": None, "curve": curve_type, "td": None}
    except ValueError as exc:
        return None, str(exc), {}


def compute_trip_time(device_row: dict, fault_a: float, element: str, role: str,
                      device_class: str, circuit: str,
                      relay_settings: dict, recloser_db: dict,
                      device_map: dict, tripsaver_db: dict) -> tuple[float | None, str, dict]:
    """
    Master dispatcher — returns (trip_time_seconds | None, note_string, settings).

    role:    'upstream'   → fuse uses Melting curve
             'downstream' → fuse uses Clearing curve
    element: 'phase' | 'ground'
    settings keys: 'pickup' (primary A), 'curve' (str), 'td' (float | None)
    """
    equip_num = device_row.get("Equipment Number", "").strip()
    equip_id  = device_row.get("Equipment ID", "").strip()
    base_kv   = _safe_float(device_row.get("Base Voltage (kVLL)"), 12.47)

    if device_class == "breaker":
        return _breaker_time(circuit, fault_a, element, relay_settings)

    if device_class == "electronic_recloser":
        return _electronic_recloser_time(equip_num, fault_a, element, recloser_db)

    if device_class == "hydraulic_recloser":
        return _hydraulic_recloser_time(equip_id, fault_a, device_map)

    if device_class == "tripsaver":
        fid = extract_fid(equip_num)
        curve_type = tripsaver_db.get(fid)
        ts_note = ""
        if curve_type is None:
            curve_type = device_map.get(equip_id, {}).get("Pickup")
        if curve_type is None:
            curve_type = "TS100T"
            ts_note = f"FID {fid} not in tripsaver DB or device map; defaulted to TS100T"
        t, err, settings = _tripsaver_time(fault_a, curve_type)
        if ts_note and not err:
            err = ts_note
        return t, err, settings

    if device_class == "fuse":
        fuse_curve = "Melting" if role == "upstream" else "Clearing"
        return _fuse_time(equip_id, fault_a, base_kv, fuse_curve)

    return None, f"Unknown device class '{device_class}'", {}


# ---------------------------------------------------------------------------
# Per-circuit coordination check
# ---------------------------------------------------------------------------

def check_circuit(circuit: str, cyme_path: Path,
                  device_map: dict, recloser_db: dict,
                  relay_settings: dict, tripsaver_db: dict) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Run all coordination checks for one circuit.
    Returns (violations, discrepancies, indeterminate). If a voltage class
    mismatch is found, the circuit is skipped and discrepancies are returned.
    """
    devices = load_cyme_report(cyme_path)

    discrepancies = _check_voltage_class(circuit, devices)
    if discrepancies:
        return [], discrepancies, []

    violations: list[dict] = []
    indeterminate: list[dict] = []

    for equip_num, d_row in devices.items():
        if not d_row.get("Upstream Protective Device", "").strip():
            continue  # top-level breaker — no upstream in scope

        d_equip_id  = d_row.get("Equipment ID", "").strip()
        d_cyme_type = d_row.get("Device Type", "").strip()

        # Skip electronic reclosers in ignored groups — not coordination targets
        if _is_ignored_recloser(equip_num, d_equip_id, d_cyme_type,
                                 device_map, recloser_db):
            continue

        u_row = _effective_upstream(equip_num, devices, device_map, recloser_db)
        if u_row is None:
            continue  # upstream on a different circuit or not found

        upstream_key = u_row.get("Equipment Number", "").strip()
        u_equip_id  = u_row.get("Equipment ID", "").strip()
        u_cyme_type = u_row.get("Device Type", "").strip()

        u_class = classify_device(u_equip_id, u_cyme_type, device_map)
        d_class = classify_device(d_equip_id, d_cyme_type, device_map)

        pair = (u_class, d_class)
        threshold = THRESHOLDS.get(pair)
        if threshold is None:
            continue  # unsupported pair (e.g. unknown device type)

        notes: list[str] = []

        # ---- Phase check -----------------------------------------------
        fault_ph = _safe_float(d_row.get("LLL / LLLG Max"))
        t_up_ph = t_dn_ph = margin_ph = None
        s_up_ph: dict = {}
        s_dn_ph: dict = {}
        ph_violation = False

        if fault_ph > 0.0:
            t_up_ph, note, s_up_ph = compute_trip_time(
                u_row, fault_ph, "phase", "upstream",
                u_class, circuit, relay_settings, recloser_db, device_map, tripsaver_db)
            if note:
                notes.append(f"Phase upstream: {note}")

            t_dn_ph, note, s_dn_ph = compute_trip_time(
                d_row, fault_ph, "phase", "downstream",
                d_class, circuit, relay_settings, recloser_db, device_map, tripsaver_db)
            if note:
                notes.append(f"Phase downstream: {note}")

            if t_up_ph is not None and t_dn_ph is not None:
                margin_ph   = t_up_ph - t_dn_ph
                ph_violation = margin_ph < threshold

        # ---- Ground check ----------------------------------------------
        fault_gnd = _safe_float(d_row.get("LG Max"))
        t_up_gnd = t_dn_gnd = margin_gnd = None
        s_up_gnd: dict = {}
        s_dn_gnd: dict = {}
        gnd_violation = False

        if fault_gnd > 0.0:
            t_up_gnd, note, s_up_gnd = compute_trip_time(
                u_row, fault_gnd, "ground", "upstream",
                u_class, circuit, relay_settings, recloser_db, device_map, tripsaver_db)
            if note:
                notes.append(f"Ground upstream: {note}")

            t_dn_gnd, note, s_dn_gnd = compute_trip_time(
                d_row, fault_gnd, "ground", "downstream",
                d_class, circuit, relay_settings, recloser_db, device_map, tripsaver_db)
            if note:
                notes.append(f"Ground downstream: {note}")

            if t_up_gnd is not None and t_dn_gnd is not None:
                margin_gnd    = t_up_gnd - t_dn_gnd
                gnd_violation = margin_gnd < threshold

        # Collect indeterminate pairs — fault current present but trip time unresolvable
        is_indeterminate = (
            (fault_ph  > 0 and (t_up_ph  is None or t_dn_ph  is None)) or
            (fault_gnd > 0 and (t_up_gnd is None or t_dn_gnd is None))
        )

        # Emit a row only when at least one element has a confirmed violation
        if not (ph_violation or gnd_violation):
            if is_indeterminate:
                phase_notes  = "; ".join(n for n in notes if n.startswith("Phase"))
                ground_notes = "; ".join(n for n in notes if n.startswith("Ground"))
                indeterminate.append({
                    "Circuit":                     circuit,
                    "Upstream_Equipment_Number":   upstream_key,
                    "Upstream_Equipment_ID":       u_equip_id,
                    "Downstream_Equipment_Number": equip_num,
                    "Downstream_Equipment_ID":     d_equip_id,
                    "Check_Type":                  f"{_CLASS_LABEL[u_class]}-{_CLASS_LABEL[d_class]}",
                    "Phase_Note":                  phase_notes,
                    "Ground_Note":                 ground_notes,
                })
            continue

        u_customers = int(_safe_float(u_row.get("Total downstream Customers")))
        d_customers = int(_safe_float(d_row.get("Total downstream Customers")))

        check_type = f"{_CLASS_LABEL[u_class]}-{_CLASS_LABEL[d_class]}"

        def _fmt_setting(val) -> str:
            if val is None:
                return ""
            return f"{val:.4f}" if isinstance(val, float) else str(val)

        violations.append({
            "Circuit":                          circuit,
            "Upstream_Equipment_Number":        upstream_key,
            "Upstream_Equipment_ID":            u_equip_id,
            "Upstream_Device_Type":             _CLASS_LABEL[u_class],
            "Downstream_Equipment_Number":      equip_num,
            "Downstream_Equipment_ID":          d_equip_id,
            "Downstream_Device_Type":           _CLASS_LABEL[d_class],
            "Check_Type":                       check_type,
            "Phase_Fault_Current_A":            fault_ph   if fault_ph   > 0 else "",
            "Phase_Upstream_Time_s":            f"{t_up_ph:.4f}"  if t_up_ph  is not None else "",
            "Phase_Downstream_Time_s":          f"{t_dn_ph:.4f}"  if t_dn_ph  is not None else "",
            "Phase_Margin_s":                   f"{margin_ph:.4f}" if margin_ph is not None else "",
            "Phase_Required_Margin_s":          threshold,
            "Phase_Violation":                  ph_violation,
            "Ground_Fault_Current_A":           fault_gnd  if fault_gnd  > 0 else "",
            "Ground_Upstream_Time_s":           f"{t_up_gnd:.4f}" if t_up_gnd  is not None else "",
            "Ground_Downstream_Time_s":         f"{t_dn_gnd:.4f}" if t_dn_gnd  is not None else "",
            "Ground_Margin_s":                  f"{margin_gnd:.4f}" if margin_gnd is not None else "",
            "Ground_Required_Margin_s":         threshold,
            "Ground_Violation":                 gnd_violation,
            "Upstream_Phase_Pickup_A":          _fmt_setting(s_up_ph.get("pickup")),
            "Upstream_Phase_Curve":             s_up_ph.get("curve", ""),
            "Upstream_Phase_TD":                _fmt_setting(s_up_ph.get("td")),
            "Upstream_Ground_Pickup_A":         _fmt_setting(s_up_gnd.get("pickup")),
            "Upstream_Ground_Curve":            s_up_gnd.get("curve", ""),
            "Upstream_Ground_TD":               _fmt_setting(s_up_gnd.get("td")),
            "Downstream_Phase_Pickup_A":        _fmt_setting(s_dn_ph.get("pickup")),
            "Downstream_Phase_Curve":           s_dn_ph.get("curve", ""),
            "Downstream_Phase_TD":              _fmt_setting(s_dn_ph.get("td")),
            "Downstream_Ground_Pickup_A":       _fmt_setting(s_dn_gnd.get("pickup")),
            "Downstream_Ground_Curve":          s_dn_gnd.get("curve", ""),
            "Downstream_Ground_TD":             _fmt_setting(s_dn_gnd.get("td")),
            "Upstream_Total_Customers":         u_customers,
            "Downstream_Total_Customers":       d_customers,
            "Customers_Erroneously_Disconnected": u_customers - d_customers,
            "Upstream_Device_Downstream_A_PH_cKVA":    u_row.get("Downstream A PH cKVA", ""),
            "Upstream_Device_Downstream_B_PH_cKVA":    u_row.get("Downstream B PH cKVA", ""),
            "Upstream_Device_Downstream_C_PH_cKVA":    u_row.get("Downstream C PH cKVA", ""),
            "Upstream_Device_Total_Downstream_3PH_cKVA": u_row.get("Total downstream 3PH cKVA", ""),
            "Downstream_Device_Downstream_A_PH_cKVA":  d_row.get("Downstream A PH cKVA", ""),
            "Downstream_Device_Downstream_B_PH_cKVA":  d_row.get("Downstream B PH cKVA", ""),
            "Downstream_Device_Downstream_C_PH_cKVA":  d_row.get("Downstream C PH cKVA", ""),
            "Downstream_Device_Total_Downstream_3PH_cKVA": d_row.get("Total downstream 3PH cKVA", ""),
            "Notes":                            "; ".join(notes),
        })

    return violations, [], indeterminate


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def write_violations(violations: list[dict], output_path: Path) -> None:
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=OUTPUT_FIELDNAMES)
        writer.writeheader()
        writer.writerows(violations)
    print(f"Wrote {len(violations)} violation(s) to {output_path}")


def write_discrepancies(discrepancies: list[dict], output_path: Path) -> None:
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["Circuit", "Base_Voltage_kVLL"])
        writer.writeheader()
        writer.writerows(discrepancies)
    print(f"Wrote {len(discrepancies)} voltage discrepancy(s) to {output_path}")


def write_indeterminate(indeterminate: list[dict], output_path: Path) -> None:
    fieldnames = [
        "Circuit",
        "Upstream_Equipment_Number",
        "Upstream_Equipment_ID",
        "Downstream_Equipment_Number",
        "Downstream_Equipment_ID",
        "Check_Type",
        "Phase_Note",
        "Ground_Note",
    ]
    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(indeterminate)
    print(f"Wrote {len(indeterminate)} indeterminate pair(s) to {output_path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    print("Loading reference data...")
    device_map     = load_device_map()
    tripsaver_db   = load_tripsaver_db()
    recloser_db    = load_recloser_db()
    relay_settings = load_relay_settings()

    print(f"  Device map:      {len(device_map):>5} entries")
    print(f"  Tripsaver DB:    {len(tripsaver_db):>5} entries")
    print(f"  Recloser DB:     {len(recloser_db):>5} entries")
    print(f"  Relay settings:  {len(relay_settings):>5} circuits")

    cyme_files = sorted(REPORTS_DIR.glob("CYME_Device_Report_*.csv"))
    print(f"  CYME reports:    {len(cyme_files):>5} circuits\n")

    all_violations: list[dict] = []
    all_discrepancies: list[dict] = []
    all_indeterminate: list[dict] = []
    for cyme_path in cyme_files:
        circuit = cyme_path.stem.rsplit("_", 1)[-1]
        violations, discrepancies, indeterminate = check_circuit(
            circuit, cyme_path, device_map, recloser_db, relay_settings, tripsaver_db)
        if discrepancies:
            print(f"  Circuit {circuit}: voltage mismatch — skipped")
        elif violations:
            print(f"  Circuit {circuit}: {len(violations)} violation(s)")
        if indeterminate:
            print(f"  Circuit {circuit}: {len(indeterminate)} indeterminate pair(s)")
        all_violations.extend(violations)
        all_discrepancies.extend(discrepancies)
        all_indeterminate.extend(indeterminate)

    print()
    write_violations(all_violations, OUTPUT_PATH)
    write_discrepancies(all_discrepancies, DISCREPANCY_PATH)
    write_indeterminate(all_indeterminate, INDETERMINATE_PATH)
    print(f"\nDone. {len(all_violations)} total violation(s) across {len(cyme_files)} circuit(s).")


if __name__ == "__main__":
    main()
