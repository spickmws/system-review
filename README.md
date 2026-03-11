# System Review — DEC Protection Coordination Tool

Bulk analysis tool that evaluates coordination margins between protective devices
(breakers, reclosers, fuses) on Duke Energy Carolinas distribution circuits.

For each upstream/downstream device pair found in a CYME Device Report, the tool
computes trip times for phase and ground faults and flags any pair that falls below
the required coordination margin.

---

## Requirements

- Python 3.10+
- `numpy`
- `openpyxl`

```
pip install numpy openpyxl
```

---

## Configuration

Two paths are configured at the top of `main.py` and can be overridden via
environment variables:

| Variable | Default | Purpose |
|---|---|---|
| `COORD_REPORTS_DIR` | `data/reports/` | Folder containing CYME Device Report and Feeder Relay Settings CSV files |
| `COORD_RECLOSER_DB` | `//nascharf06/DPAC/21 Recloser Settings Database/Recloser Database/RecloserDatabase.xlsx` | Electronic recloser settings database |

---

## Input Files

### CYME Device Reports
Filename pattern: `CYME_Device_Report_<circuit>.csv`

One file per circuit, placed in `COORD_REPORTS_DIR`. Each row represents a
protective device (`Breaker`, `Fuse`, or `Recloser`) with status
`ConnectedClose`. Key columns used:

| Column | Description |
|---|---|
| `Equipment Number` | Unique device identifier on this circuit |
| `Equipment ID` | Hardware model key (looked up in `protective_device_mapping.csv`) |
| `Device Type` | `Breaker`, `Fuse`, or `Recloser` |
| `Upstream Protective Device` | Equipment Number of the next upstream device |
| `LLL / LLLG Max` | Three-phase fault current (A) at this location |
| `LG Max` | Line-to-ground fault current (A) at this location |
| `Base Voltage (kVLL)` | Nominal line-to-line voltage (used for fuse curve selection) |
| `Total downstream Customers` | Customer count served through this device |

### Feeder Relay Settings
Filename pattern: `Feeder_Relay_Settings_<circuit>.csv`

Key-value row format. Required keys for breaker trip-time calculation:

| Key | Description |
|---|---|
| `CTR` | Current transformer ratio |
| `51P1P` | Phase overcurrent pickup (secondary amps) |
| `51P1C` | Phase overcurrent curve (U1–U5) |
| `51P1TD` | Phase time dial |
| `51G1P` | Ground overcurrent pickup (secondary amps) |
| `51G1C` | Ground overcurrent curve |
| `51G1TD` | Ground time dial |

### RecloserDatabase.xlsx
Sheet: `Reclosers`. One row per electronic recloser, keyed by `FID`.

| Column | Description |
|---|---|
| `FID` | Feeder ID (integer); matched against the last `_`-delimited token of Equipment Number |
| `Group` | Operational group number; reclosers in groups **5, 7, or 8** are excluded from coordination |
| `CTR` | Current transformer ratio |
| `ph` / `gPU` | Phase / ground pickup (secondary amps) |
| `phSC` / `gSC` | Phase / ground slow curve (`U1`–`U5` or numeric SEL code) |
| `phSTD` / `gSTD` | Phase / ground time dial |

### protective_device_mapping.csv
Located at `data/protective_device_mapping.csv`. Maps Equipment ID to device class
and pickup current.

| Pickup value | Device class |
|---|---|
| `Electronic` | Electronic recloser |
| `TS` | TripSaver sectionalizer |
| Numeric (e.g. `200`) | Hydraulic recloser — value is pickup in primary amps |
| *(blank)* | Breaker or Fuse (determined by Device Type column) |

---

## How It Works

### Device Classification
Each device is classified as one of:
`breaker` · `electronic_recloser` · `hydraulic_recloser` · `tripsaver` · `fuse` · `unknown`

### Ignored Reclosers
Electronic reclosers assigned to **groups 5, 7, or 8** are skipped entirely.
When such a recloser sits between two coordination points, the tool looks through
it to find the next eligible upstream device.

### Coordination Margins
Required time margins (seconds) between device classes:

| Upstream | Downstream | Margin |
|---|---|---|
| Breaker | Recloser / TripSaver | 0.20 s |
| Recloser / TripSaver | Recloser / TripSaver | 0.22 s |
| Any | Fuse | 0.00 s |

### Trip-Time Calculation

| Device | Curve source | Notes |
|---|---|---|
| Breaker | U-curve equations (U1–U5) | Settings from Feeder Relay Settings CSV |
| Electronic recloser | U-curve or SEL numeric curve | Settings from RecloserDatabase.xlsx |
| Hydraulic recloser | `data/hydraulic_i_t.csv` | Slow curve only; `i = fault_A / pickup` |
| TripSaver | `data/ts_i_t.csv` | All units use `TS100T` curve |
| Fuse | `data/fuses_i_t.csv` | Melting (upstream role) or Clearing (downstream role) |

All CSV curves are stored as log10(i) vs log10(t) and interpolated in log-log space.

---

## Output

Results are written to `violations_report.csv` in the project root. A row is
emitted for every device pair where the computed margin is less than the required
threshold for either phase or ground faults.

Key output columns include trip times, margins, violation flags, protection
settings (pickup, curve, time dial) for both devices, and downstream customer
counts.

---

## Usage

```
python main.py
```

Override paths if needed:

```
set COORD_REPORTS_DIR=C:\path\to\reports
set COORD_RECLOSER_DB=C:\path\to\RecloserDatabase.xlsx
python main.py
```

---

## Utilities

### `test_hydraulic_times.py`
Prints a table of slow-curve opening times for every hydraulic recloser in
`protective_device_mapping.csv` at configurable fault levels.

```
python test_hydraulic_times.py
```

Edit the `FAULT_LEVELS` list near the top of the file to change the fault
current values tested.

---

## Project Structure

```
system-review/
├── main.py                            # Main coordination engine
├── test_hydraulic_times.py            # Hydraulic recloser opening-time table
├── violations_report.csv              # Output (generated on each run)
├── data/
│   ├── protective_device_mapping.csv  # Equipment ID → device class + pickup
│   ├── curves_i_t.csv                 # Relay/overcurrent curves
│   ├── fuses_i_t.csv                  # Fuse melting/clearing curves
│   ├── hydraulic_i_t.csv              # Hydraulic recloser fast/slow curves
│   ├── ts_i_t.csv                     # TripSaver curves
│   ├── u_curves.txt                   # ANSI/IEEE inverse-time equations
│   ├── reports/                       # Place CYME and relay settings CSVs here
│   └── samples/
│       └── RecloserDatabase_sample.xlsx
└── src/
    ├── protective_device_coordination.py  # Trip-time lookup engine (get_trip_time)
    ├── CPAT_Report.py                     # Legacy CYME integration reference
    └── extract_equipment.py               # Equipment ID extraction utility
```
