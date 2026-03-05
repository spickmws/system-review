from protective_device_coordination import get_trip_time

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