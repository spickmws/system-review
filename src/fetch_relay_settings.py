import csv
import re
import urllib.request
from xml.etree import ElementTree as ET
from pathlib import Path

BASE_URL    = "http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp"
LIST_PATH   = Path(__file__).parent.parent / "data" / "references" / "station_list.csv"
OUTPUT_PATH = Path(__file__).parent.parent / "relay_settings.csv"

# ── copied verbatim from CPAT_Report_Batch.py lines 49–204 ──────────────────
relay_settings = {
    'SEL-351S-6':   ['RID','CTR','E79','51P1P','51P1C','51P1TD','51G1P','51G1C','51G1TD',
                     '50P1P','50P2P','50P3P','50G1P','50G2P','50G3P',
                     '67P1D','67P2D','67P3D','67G1D','67G2D','67G3D'],
    'SEL-351S-7':   ['RID','CTR','E79','51P1P','51P1C','51P1TD','51G1P','51G1C','51G1TD',
                     '50P1P','50P2P','50P3P','50G1P','50G2P','50G3P',
                     '67P1D','67P2D','67P3D','67G1D','67G2D','67G3D'],
    'SEL-351-6-V0': ['RID','CTR','E79','51PP','51PC','51PTD','51GP','51GC','51GTD',
                     '50P1P','50P4P','50G1P','50G3P','67P1D','67P4D','67G1D','67G3D'],
    'SEL-351-7-V0': ['RID','CTR','E79','51PP','51PC','51PTD','51GP','51GC','51GTD',
                     '50P1P','50P4P','50G1P','50G3P','67P1D','67P4D','67G1D','67G3D'],
    'SEL-351':      ['RID','CTR','CTRN','E79','51PP','51PC','51PTD','51NP','51NC','51NTD',
                     '50P1P','50P2P','50P3P','50N1P','50N2P','50N3P',
                     '67P1D','67P2D','67P3D','67N1D','67N2D','67N3D'],
    'SEL-351R-2':   ['RID','CTR','E79','51P1P','51P1C','51P1TD','51G1P','51G1C','51G1TD',
                     '50P2P','67P2D','50G2P','67G2D'],
    'SEL-351R-4':   ['RID','CTR','E79','51P1P','51P1C','51P1TD','51G1P','51G1C','51G1TD',
                     '50P2P','67P2D','50G2P','67G2D'],
}
other_relays = ['IAC51', 'IAC53', 'CO-8']
# ────────────────────────────────────────────────────────────────────────────

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
        {"id": int(tryfloat(c.attrib["STANUM"])),
         "name": c.attrib["S01"],
         "aspen": c.attrib["ID"].strip()}
        for c in tree
    ]

def parse_cb_list(eq_tree: ET.Element) -> list[str]:
    """Reproduce lines 1009–1016 of CPAT_Report_Batch.py."""
    texts = [e.text for e in eq_tree if e.text]
    cb_list = [t for t in texts if t.isnumeric() and len(t) == 4]
    cb_list += [t for t in texts if re.search(r'^OCR\d+ \([0-9]{4}\)', t)]
    cb_list += [t for t in texts if re.search(r'^\d+ \([0-9]{4}\)', t)]
    cb_list += [t for t in texts if re.search(r'^[0-9]{4} \(\d+\)', t)]
    cb_list += [t for t in texts if re.search(r'^[0-9]{4} \(OCR\d+\)', t)]
    cb_list += [t for t in texts if re.search(r'^[0-9]{4} \(NEW\)', t)]
    cb_list += [t for t in texts if len(t) == 4 and re.search(r'^[0-9]{2}AX', t)]
    cb_list += [t for t in texts if re.search(r'^[0-9]{4} \[NEW\]', t)]
    return cb_list

def cb_num_from(cb: str) -> str:
    """Reproduce lines 1020–1026 of CPAT_Report_Batch.py."""
    if len(cb) <= 4:
        return cb
    if cb[:4].isnumeric() and cb[4] == ' ':
        return cb[:4]
    return cb[-5:-1]

def fetch_relay_settings_for_station(sta: dict) -> list[dict]:
    rows = []
    aspen = sta["aspen"]

    eq_tree = get(f"{BASE_URL}?a=2&lid={aspen}&c=DEC")
    if not list(eq_tree):
        eq_tree = get(f"{BASE_URL}?a=2&lid={aspen}%20&c=DEC")

    for cb in parse_cb_list(eq_tree):
        cb_num = cb_num_from(cb)
        breaker_tree = get(
            f"{BASE_URL}?a=4&lid={aspen}&cid={cb.replace(' ', '%20')}&c=DEC"
        )
        rels = breaker_tree.findall("RELAY")

        relay_id, relay_type = "", ""
        settings = {}

        for rel in rels:
            relay_id   = rel.attrib["RELAYID"]
            relay_type = rel.attrib["RELAYTYPE"]

            if relay_type in relay_settings:
                settings_tree = get(f"{BASE_URL}?a=5&rid={relay_id}&c=DEC")
                try:
                    group = settings_tree.find("SETTINGS").find("REQUEST").find("GROUPS")[0]
                except (TypeError, IndexError):
                    try:
                        group = settings_tree.find("SETTINGS").find("REQUEST").find("GROUPS")[5]
                    except (TypeError, IndexError):
                        break  # no parseable group; relay recorded below with empty settings

                wanted = set(relay_settings[relay_type])
                for elem in group:
                    a = elem.attrib
                    if a.get("SETTINGNAME") in wanted:
                        settings[a["SETTINGNAME"]] = a.get("SETTING", "")
                break  # stop after first SEL relay

            elif rel.attrib.get("S01") == "79":
                continue
            elif any(abbr in relay_type for abbr in other_relays):
                for abbr in other_relays:
                    if abbr in relay_type:
                        relay_type = abbr
                break
            else:
                continue

        base = {
            "Station_ID":   sta["id"],
            "Station_Name": sta["name"],
            "Aspen_ID":     aspen,
            "Breaker_ID":   cb_num,
            "Relay_ID":     relay_id,
            "Relay_Type":   relay_type,
        }
        if settings:
            for name, val in settings.items():
                rows.append({**base, "Setting_Name": name, "Setting_Value": val})
        else:
            rows.append({**base, "Setting_Name": "", "Setting_Value": ""})

    return rows

def main():
    circuit_ids = set()
    with open(LIST_PATH, newline="", encoding="utf-8") as f:
        for line in f:
            cid = line.strip()
            if cid:
                circuit_ids.add(cid)

    print(f"Loaded {len(circuit_ids)} circuit IDs from {LIST_PATH}")
    print("Fetching station list …")
    all_stations = fetch_all_stations()

    in_scope = [
        sta for sta in all_stations
        if any(
            cid.startswith(str(sta["id"])) and cid[len(str(sta["id"])):].isdigit()
            for cid in circuit_ids
        )
    ]
    print(f"{len(in_scope)} stations in scope (of {len(all_stations)} total)")

    fieldnames = ["Station_ID","Station_Name","Aspen_ID","Breaker_ID",
                  "Relay_ID","Relay_Type","Setting_Name","Setting_Value"]
    total_rows = 0

    with open(OUTPUT_PATH, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for i, sta in enumerate(in_scope, 1):
            print(f"  [{i}/{len(in_scope)}] {sta['name']} (STANUM {sta['id']}) …", end=" ", flush=True)
            try:
                rows = fetch_relay_settings_for_station(sta)
                writer.writerows(rows)
                total_rows += len(rows)
                print(f"{len(rows)} rows")
            except Exception as e:
                print(f"ERROR: {e}")

    print(f"\nDone. Wrote {total_rows} row(s) to {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
