"""
protective_device_coordination.py
==================================
Unified function for querying trip/operating times from protective device
coordination data.

Supported device types
----------------------
1. Protection curves  (curves_i_t.csv)   – relay/overcurrent curves stored in
                                            log10-log10 space.
2. U-curves           (u_curves.txt)      – ANSI/IEEE inverse-time equations
                                            (U1–U5) that require a time-dial.
3. TripSaver (TS)     (ts_i_t.csv)       – sectionalizer curves stored in
                                            log10-log10 space.
4. Fuses              (fuses_i_t.csv)     – melting / clearing curves stored in
                                            log10-log10 space.
5. Hydraulic recloser (hydraulic_i_t.csv)– fast / slow curves stored in
                                            log10-log10 space.

All CSV data are stored as log10(i) vs log10(t).  Interpolation is performed
in log-log space and the result is converted back to linear time (seconds).

Usage examples
--------------
from protective_device_coordination import get_trip_time

# 1. Protection curve (curves_i_t) – just needs curve type and i (linear amps
#    or multiples of pickup, whatever units the original data use)
t = get_trip_time(device='curve', curve_type='101', i=2.5)

# 2. U-curve – needs curve name, i (multiple of pickup), and time dial
t = get_trip_time(device='ucurve', curve_type='U2', i=5.0, time_dial=0.5)

# 3. TripSaver – needs TS type and i
t = get_trip_time(device='ts', curve_type='TS80T', i=400)

# 4. Fuse – needs fuse type, curve (melting or clearing), and i
t = get_trip_time(device='fuse', curve_type='100T', curve='Melting', i=200)

# 5. Hydraulic recloser – needs recloser type, curve (fast or slow), and i
t = get_trip_time(device='hydraulic', curve_type='L', curve='fast', i=100)
"""

import os
import csv
import math
import numpy as np
from pathlib import Path

# ---------------------------------------------------------------------------
# Locate data files relative to this script so it works regardless of the
# current working directory.
# ---------------------------------------------------------------------------
_HERE = Path(__file__).parent

# Allow the caller to override the data directory via an environment variable.
_DATA_DIR = Path(os.environ.get("COORD_DATA_DIR", str(_HERE)))

_CURVES_FILE    = _DATA_DIR / "curves_i_t.csv"
_FUSES_FILE     = _DATA_DIR / "fuses_i_t.csv"
_HYDRAULIC_FILE = _DATA_DIR / "hydraulic_i_t.csv"
_TS_FILE        = _DATA_DIR / "ts_i_t.csv"

# ---------------------------------------------------------------------------
# U-curve equations  (i = multiple of pickup, already linear)
# ---------------------------------------------------------------------------
_U_CURVES = {
    "U1": lambda i, td: td * (0.0226   + (0.0104   / (i ** 0.02 - 1))),
    "U2": lambda i, td: td * (0.180    + (5.95     / (i ** 2   - 1))),
    "U3": lambda i, td: td * (0.0963   + (3.88     / (i ** 2   - 1))),
    "U4": lambda i, td: td * (0.02434  + (5.64     / (i ** 2   - 1))),
    "U5": lambda i, td: td * (0.00262  + (0.00342  / (i ** 0.02 - 1))),
}

# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _parse_cell(cell: str) -> list[float]:
    """Split a semicolon-delimited cell into a list of floats."""
    return [float(v) for v in cell.strip().split(";")]


def _load_csv(filepath: Path) -> list[dict]:
    """Read a CSV and return a list of row dicts."""
    with open(filepath, "r", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def _loglog_interp(log_i_data: list[float], log_t_data: list[float],
                   i_linear: float) -> float:
    """
    Interpolate in log-log space and return t in linear seconds.

    Parameters
    ----------
    log_i_data : list of log10(i) values (already in log space)
    log_t_data : list of log10(t) values (already in log space)
    i_linear   : query current in the same linear units as the original data

    Returns
    -------
    t in seconds (linear)

    Raises
    ------
    ValueError  if i_linear is outside the range of the data
    """
    log_i_query = math.log10(i_linear)

    log_i_arr = np.array(log_i_data)
    log_t_arr = np.array(log_t_data)

    if log_i_query < log_i_arr[0] or log_i_query > log_i_arr[-1]:
        raise ValueError(
            f"i={i_linear} (log10={log_i_query:.4f}) is outside the "
            f"data range [10^{log_i_arr[0]:.4f}, 10^{log_i_arr[-1]:.4f}]."
        )

    log_t_query = np.interp(log_i_query, log_i_arr, log_t_arr)
    return 10 ** log_t_query


# ---------------------------------------------------------------------------
# Per-source lookup functions
# ---------------------------------------------------------------------------

def _get_curve_time(curve_type: str, i: float) -> float:
    """Look up time from curves_i_t.csv."""
    rows = _load_csv(_CURVES_FILE)
    for row in rows:
        if row["type"].strip() == str(curve_type).strip():
            log_i = _parse_cell(row["i"])
            log_t = _parse_cell(row["t"])
            return _loglog_interp(log_i, log_t, i)
    available = sorted({r["type"].strip() for r in rows})
    raise ValueError(
        f"Curve type '{curve_type}' not found in curves_i_t.csv.\n"
        f"Available types: {available}"
    )


def _get_ucurve_time(curve_type: str, i: float, time_dial: float) -> float:
    """Evaluate a U-curve equation."""
    key = curve_type.upper().strip()
    if key not in _U_CURVES:
        raise ValueError(
            f"U-curve '{curve_type}' not recognised. "
            f"Valid options: {sorted(_U_CURVES.keys())}"
        )
    if i <= 1.0:
        raise ValueError(
            f"i must be > 1.0 for U-curves (got {i}). "
            "i is the multiple of pickup current."
        )
    return _U_CURVES[key](i, time_dial)


def _get_ts_time(curve_type: str, i: float) -> float:
    """Look up time from ts_i_t.csv."""
    rows = _load_csv(_TS_FILE)
    for row in rows:
        if row["type"].strip().upper() == str(curve_type).strip().upper():
            log_i = _parse_cell(row["i"])
            log_t = _parse_cell(row["t"])
            return _loglog_interp(log_i, log_t, i)
    available = sorted({r["type"].strip() for r in rows})
    raise ValueError(
        f"TripSaver type '{curve_type}' not found in ts_i_t.csv.\n"
        f"Available types: {available}"
    )


def _get_fuse_time(curve_type: str, curve: str, i: float) -> float:
    """Look up time from fuses_i_t.csv.

    Parameters
    ----------
    curve_type : fuse rating/designation, e.g. '100T', '15E_14.4kV'
    curve      : 'Melting' or 'Clearing' (case-insensitive)
    i          : current (linear, same units as the source data)
    """
    curve_norm = curve.strip().capitalize()          # 'melting' → 'Melting'
    if curve_norm not in ("Melting", "Clearing"):
        raise ValueError(
            f"curve must be 'Melting' or 'Clearing' (got '{curve}')."
        )
    rows = _load_csv(_FUSES_FILE)
    for row in rows:
        if (row["type"].strip() == str(curve_type).strip() and
                row["curve"].strip().capitalize() == curve_norm):
            log_i = _parse_cell(row["i"])
            log_t = _parse_cell(row["t"])
            return _loglog_interp(log_i, log_t, i)
    available = sorted({r["type"].strip() for r in rows})
    raise ValueError(
        f"Fuse type '{curve_type}' / curve '{curve}' not found in "
        f"fuses_i_t.csv.\nAvailable fuse types: {available}"
    )


def _get_hydraulic_time(curve_type: str, curve: str, i: float) -> float:
    """Look up time from hydraulic_i_t.csv.

    Parameters
    ----------
    curve_type : recloser type, e.g. 'L', 'V4L_V4E', '4E'
    curve      : 'fast' or 'slow' (case-insensitive)
    i          : current (linear, same units as the source data)
    """
    curve_norm = curve.strip().lower()
    if curve_norm not in ("fast", "slow"):
        raise ValueError(
            f"curve must be 'fast' or 'slow' (got '{curve}')."
        )
    rows = _load_csv(_HYDRAULIC_FILE)
    for row in rows:
        if (row["type"].strip() == str(curve_type).strip() and
                row["curve"].strip().lower() == curve_norm):
            log_i = _parse_cell(row["i"])
            log_t = _parse_cell(row["t"])
            return _loglog_interp(log_i, log_t, i)
    available = sorted({r["type"].strip() for r in rows})
    raise ValueError(
        f"Hydraulic type '{curve_type}' / curve '{curve}' not found in "
        f"hydraulic_i_t.csv.\nAvailable types: {available}"
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_trip_time(
    device: str,
    curve_type: str,
    i: float,
    *,
    curve: str | None = None,
    time_dial: float | None = None,
) -> float:
    """
    Return the trip/operating time (seconds) for a protective device.

    Parameters
    ----------
    device : str
        Device category.  One of:
          'curve'     – protection curve from curves_i_t.csv
          'ucurve'    – ANSI/IEEE inverse-time equation (U1–U5)
          'ts'        – TripSaver from ts_i_t.csv
          'fuse'      – fuse from fuses_i_t.csv
          'hydraulic' – hydraulic recloser from hydraulic_i_t.csv

    curve_type : str
        Identifier for the specific device/curve:
          'curve'     → numeric type code, e.g. '101', '200'
          'ucurve'    → 'U1' … 'U5'
          'ts'        → 'TS80T', 'TS100T', 'TS150E'
          'fuse'      → e.g. '100T', '15E_14.4kV', '50K'
          'hydraulic' → e.g. 'L', 'V4L_V4E', '4E', 'D'

    i : float
        Current in **linear units** for all device types.
          - For CSV-based devices ('curve', 'ts', 'fuse', 'hydraulic'): pass
            the actual linear current (amps, or whatever unit the original
            data use).  The function converts to log10 internally before
            interpolating.
          - For 'ucurve': pass the **multiple of pickup** as a plain linear
            ratio (e.g. 5.0 means 5× pickup).  Must be > 1.0.

    curve : str, optional
        Required for 'fuse' ('Melting' or 'Clearing', case-insensitive) and
        'hydraulic' ('fast' or 'slow', case-insensitive).

    time_dial : float, optional
        Required for 'ucurve'.  Typical range 0.5–10.

    Returns
    -------
    float
        Trip/operating time in seconds.

    Raises
    ------
    ValueError
        If the requested device/type/curve is not found, or if i is out of
        the interpolation range, or if required parameters are missing.

    Examples
    --------
    >>> get_trip_time(device='curve', curve_type='101', i=2.5)
    >>> get_trip_time(device='ucurve', curve_type='U2', i=5.0, time_dial=0.5)
    >>> get_trip_time(device='ts', curve_type='TS80T', i=400)
    >>> get_trip_time(device='fuse', curve_type='100T', curve='Melting', i=200)
    >>> get_trip_time(device='hydraulic', curve_type='L', curve='fast', i=100)
    """
    device_norm = device.strip().lower()

    if device_norm == "curve":
        return _get_curve_time(curve_type, i)

    elif device_norm == "ucurve":
        if time_dial is None:
            raise ValueError("time_dial is required for U-curve lookups.")
        return _get_ucurve_time(curve_type, i, time_dial)

    elif device_norm == "ts":
        return _get_ts_time(curve_type, i)

    elif device_norm == "fuse":
        if curve is None:
            raise ValueError(
                "curve ('Melting' or 'Clearing') is required for fuse lookups."
            )
        return _get_fuse_time(curve_type, curve, i)

    elif device_norm == "hydraulic":
        if curve is None:
            raise ValueError(
                "curve ('fast' or 'slow') is required for hydraulic lookups."
            )
        return _get_hydraulic_time(curve_type, curve, i)

    else:
        raise ValueError(
            f"Unknown device '{device}'. "
            "Valid options: 'curve', 'ucurve', 'ts', 'fuse', 'hydraulic'."
        )


# ---------------------------------------------------------------------------
# Quick self-test (run with:  python protective_device_coordination.py)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    tests = [
        dict(label="Curve 101,        i=2.0",
             kwargs=dict(device="curve", curve_type="101", i=2)),
        dict(label="U2 curve (TD=0.5), i=5.0× pickup",
             kwargs=dict(device="ucurve", curve_type="U2", i=100, time_dial=0.5)),
        dict(label="TripSaver TS80T,  i=10**2 A",
             kwargs=dict(device="ts", curve_type="TS80T", i=10**3)),
        dict(label="Fuse 100T Melting,  i=10**3 A",
             kwargs=dict(device="fuse", curve_type="100T", curve="Melting", i=2500)),
        dict(label="Fuse 100T Clearing, i=10**3 A",
             kwargs=dict(device="fuse", curve_type="100T", curve="Clearing", i=2500)),
        dict(label="Hydraulic L fast, i=2",
             kwargs=dict(device="hydraulic", curve_type="L", curve="fast", i=2)),
        dict(label="Hydraulic L slow, i=2",
             kwargs=dict(device="hydraulic", curve_type="L", curve="slow", i=2)),
    ]

    print(f"{'Test':<30} {'t (s)':>12}")
    print("-" * 44)
    for t in tests:
        try:
            result = get_trip_time(**t["kwargs"])
            print(f"{t['label']:<30} {result:>12.4f}")
        except ValueError as exc:
            print(f"{t['label']:<30} ERROR: {exc}")
