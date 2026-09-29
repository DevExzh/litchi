"""Check the reader against all retained baseline qualifications before capture."""
import analyze as a
import audit as independent
import driver as d


def main():
    builds = a.load_build("before")
    a.require_plan()
    a.require_inputs()
    a.require_stage("qualification-before")
    rows = []
    for logical in ("native", "observer"):
        for case in a.EXPECTED_CASES:
            path = d.P / "qualification" / f"before-{logical}-{case['id']}.json"
            independent.load_report(path, case, logical, 1, 0, builds["binaries"][logical],
                f"qualification/before/{logical}/{case['id']}")
            rows.append(a.load_report(path, case, logical, 1, 0, builds["binaries"][logical],
                f"qualification/before/{logical}/{case['id']}"))
    for case in a.EXPECTED_CASES:
        keys = {a.oracle_key(row) for row in rows if row["id"] == case["id"]}
        assert len(keys) == 1
    print("0832 baseline reader preflight PASS: both readers, 18 retained reports; no workloads executed")


if __name__ == "__main__":
    main()

