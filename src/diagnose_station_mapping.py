"""
diagnose_station_mapping.py

Diagnostic script to verify that circuit IDs in station_list.csv map correctly
to stations returned by the relay info API.

For each circuit ID, reports:
  - The station prefix and circuit number extracted
  - The matched API station (STANUM, raw STANUM, S01 name, Aspen ID)
  - Whether multiple stations could ambiguously match the same circuit ID
  - For out-of-scope circuit IDs, the specific reason
"""

import csv
import urllib.request
from xml.etree import ElementTree as ET
from pathlib import Path

BASE_URL    = "http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp"
LIST_PATH   = Path(__file__).parent.parent / "data" / "references" / "station_list.csv"
OUTPUT_PATH = Path(__file__).parent.parent / "station_mapping_diagnostic.csv"


def tryfloat(s):
    try:
        return float(s)
    except Exception:
        return 0.0


def get(url: str) -> ET.Element:
    return ET.fromstring(urllib.request.urlopen(url).read())


def fetch_all_stations() -> list[dict]:
    tree = get(f"{BASE_URL}?a=1&c=DEC")
    return [
        {
            "stanum":     int(tryfloat(c.attrib["STANUM"])),
            "raw_stanum": c.attrib["STANUM"],           # un-converted value from API
            "name":       c.attrib["S01"],
            "aspen":      c.attrib["ID"].strip(),
        }
        for c in tree
    ]


def find_all_matches(cid: str, all_stations: list[dict]) -> list[dict]:
    """
    Return every station whose STANUM string is a prefix of cid
    AND whose remaining suffix is all digits.
    Multiple results indicates an ambiguous match.
    """
    matches = []
    for sta in all_stations:
        prefix = str(sta["stanum"])
        suffix = cid[len(prefix):]
        if cid.startswith(prefix) and suffix.isdigit() and len(suffix) > 0:
            matches.append({**sta, "prefix": prefix, "circuit_number": suffix})
    return matches


def find_prefix_only_matches(cid: str, all_stations: list[dict]) -> list[str]:
    """
    Return stations whose STANUM is a prefix of cid but the suffix is NOT all digits.
    Helps identify near-misses where the station prefix is correct but the circuit
    portion contains unexpected characters.
    """
    near = []
    for sta in all_stations:
        prefix = str(sta["stanum"])
        suffix = cid[len(prefix):]
        if cid.startswith(prefix) and suffix and not suffix.isdigit():
            near.append(f"{sta['stanum']} \"{sta['name']}\" (suffix={suffix!r})")
    return near


def main():
    circuit_ids = []
    with open(LIST_PATH, newline="", encoding="utf-8") as f:
        for line in f:
            cid = line.strip()
            if cid:
                circuit_ids.append(cid)

    print(f"Loaded {len(circuit_ids)} circuit IDs from {LIST_PATH}")
    print("Fetching station list from API …")
    all_stations = fetch_all_stations()
    print(f"Received {len(all_stations)} stations from API\n")

    fieldnames = [
        "Circuit_ID",
        "In_Scope",
        "Out_of_Scope_Reason",
        "Circuit_Station_Prefix",
        "Circuit_Number",
        "STANUM",
        "Raw_STANUM",
        "Station_Name",
        "Aspen_ID",
        "Ambiguous_Match",
        "All_Possible_Matches",
    ]

    rows = []
    for cid in circuit_ids:
        row = dict.fromkeys(fieldnames, "")
        row["Circuit_ID"] = cid

        matches = find_all_matches(cid, all_stations)

        if not cid.strip():
            row["In_Scope"] = "No"
            row["Out_of_Scope_Reason"] = "Empty circuit ID"

        elif len(matches) == 0:
            row["In_Scope"] = "No"

            # Check for prefix-only near-misses (correct station, bad suffix)
            near = find_prefix_only_matches(cid, all_stations)
            if near:
                row["Out_of_Scope_Reason"] = (
                    "STANUM prefix found but circuit suffix is non-numeric"
                )
                row["All_Possible_Matches"] = "; ".join(near)
            else:
                # Try all possible split points to find candidate STANUMs
                candidates = []
                for split in range(1, len(cid)):
                    prefix_str = cid[:split]
                    suffix_str = cid[split:]
                    if suffix_str.isdigit():
                        try:
                            prefix_int = int(prefix_str)
                        except ValueError:
                            continue
                        for sta in all_stations:
                            if sta["stanum"] == prefix_int:
                                candidates.append(
                                    f"STANUM={prefix_int} \"{sta['name']}\" "
                                    f"(split {split}+{len(cid)-split}: "
                                    f"prefix={prefix_str}, circuit={suffix_str})"
                                )
                if candidates:
                    row["Out_of_Scope_Reason"] = (
                        "No match using current prefix logic; "
                        "possible split-point mismatch — see All_Possible_Matches"
                    )
                    row["All_Possible_Matches"] = "; ".join(candidates)
                else:
                    row["Out_of_Scope_Reason"] = (
                        "No API station STANUM matches any numeric prefix of this circuit ID"
                    )

        else:
            # Use first match (same as fetch_relay_settings.py behaviour)
            m = matches[0]
            row["In_Scope"]               = "Yes"
            row["Circuit_Station_Prefix"] = m["prefix"]
            row["Circuit_Number"]         = m["circuit_number"]
            row["STANUM"]                 = m["stanum"]
            row["Raw_STANUM"]             = m["raw_stanum"]
            row["Station_Name"]           = m["name"]
            row["Aspen_ID"]               = m["aspen"]

            if len(matches) > 1:
                row["Ambiguous_Match"] = "Yes"
                row["All_Possible_Matches"] = "; ".join(
                    f"STANUM={m2['stanum']} \"{m2['name']}\" "
                    f"(prefix={m2['prefix']}, circuit={m2['circuit_number']})"
                    for m2 in matches
                )
            else:
                row["Ambiguous_Match"] = "No"

        rows.append(row)

    with open(OUTPUT_PATH, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    in_scope   = sum(1 for r in rows if r["In_Scope"] == "Yes")
    out_scope  = sum(1 for r in rows if r["In_Scope"] == "No")
    ambiguous  = sum(1 for r in rows if r.get("Ambiguous_Match") == "Yes")

    print(f"Results: {in_scope} in scope, {out_scope} out of scope, {ambiguous} ambiguous")
    print(f"Written to {OUTPUT_PATH}")

    if out_scope:
        print("\nOut-of-scope circuit IDs:")
        for r in rows:
            if r["In_Scope"] == "No":
                print(f"  {r['Circuit_ID']}: {r['Out_of_Scope_Reason']}")
                if r["All_Possible_Matches"]:
                    print(f"    → {r['All_Possible_Matches']}")


if __name__ == "__main__":
    main()
