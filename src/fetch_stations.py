import csv
import urllib.request
from xml.etree import ElementTree as ET
from pathlib import Path

STATIONS_URL = "http://sysprot.duke-energy.com/relay_info/decv2/aspen.asp?a=1&c=DEC"
OUTPUT_PATH = Path(__file__).parent.parent / "stations.csv"


def tryfloat(s):
    try:
        return float(s)
    except Exception:
        return 0.0


def fetch_stations(url: str) -> list[dict]:
    response = urllib.request.urlopen(url).read()
    tree = ET.fromstring(response)
    stations = []
    for child in tree:
        stations.append({
            "Station_ID":   int(tryfloat(child.attrib["STANUM"])),
            "Station_Name": child.attrib["S01"],
            "Aspen_ID":     child.attrib["ID"].strip(),
            "Size_MVA":     tryfloat(child.attrib["S04"]),
        })
    return stations


def main():
    print(f"Fetching stations from {STATIONS_URL} ...")
    try:
        stations = fetch_stations(STATIONS_URL)
    except Exception as e:
        print(f"Error fetching stations: {e}")
        return
    with open(OUTPUT_PATH, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["Station_ID", "Station_Name", "Aspen_ID", "Size_MVA"])
        writer.writeheader()
        writer.writerows(stations)
    print(f"Wrote {len(stations)} station(s) to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
