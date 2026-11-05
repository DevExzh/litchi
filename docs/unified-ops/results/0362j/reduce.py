"""Reduce every retained c0 PPT attempt without selecting favorable blocks."""

import argparse
import hashlib
import json
from pathlib import Path
import statistics


def artifact(path):
    path = path.resolve()
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
    }


def save(path, value):
    path.write_text(json.dumps(value, indent=2) + "\n")


def checked_artifact(expected):
    actual = artifact(Path(expected["path"]))
    assert actual == expected, (actual, expected)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--coord", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    coord = args.coord.resolve()
    args.out.mkdir(parents=True, exist_ok=True)
    runs = []
    comparisons = {}
    for group in ("sip", "original"):
        blocks = []
        identities = set()
        for directory in sorted(coord.glob(f"ppt-phase-c0-{group}-v*")):
            if not directory.is_dir():
                continue
            source = directory / "measurements.json"
            if not source.exists():
                continue
            raw = json.loads(source.read_text())
            assert raw["after_commit"] == "c0f346607bb28559ae21fa780c835dc70e87fb0b"
            for item in (*raw["binaries"].values(), raw["fixture"], raw["harness"], raw["runner"]):
                checked_artifact(item)
            identities.add(tuple(raw["binaries"][s]["sha256"] for s in ("before", "after")) + (
                raw["fixture"]["sha256"], raw["harness"]["sha256"], raw["runner"]["sha256"],
            ))
            hosts = list(directory.glob("*.host.json"))
            dirty_hosts = [artifact(p) for p in sorted(hosts) if json.loads(p.read_text())["blockers"]]
            inventory = [artifact(p) for p in sorted(directory.rglob("*")) if p.is_file()]
            inventory_path = args.out / f"{directory.name}-artifacts.json"
            save(inventory_path, inventory)
            for block in raw["paired"]:
                assert len(block["legs"]) == 4
                for leg in block["legs"]:
                    assert min(leg["spans_ns"].values()) >= raw["minimum_phase_ns"]
                    checked_artifact(leg["receipt"])
                    receipt = json.loads(Path(leg["receipt"]["path"]).read_text())
                    assert receipt["exit_code"] == 0 and receipt["overlap"] is None
                    for item in receipt["artifacts"]:
                        checked_artifact(item)
                    leg_hosts = list(directory.glob(leg["name"] + "-*.host.json"))
                    assert len(leg_hosts) >= 2
                    assert all(not json.loads(p.read_text())["blockers"] for p in leg_hosts)
                blocks.append(block)
            runs.append({
                "name": directory.name,
                "measurement_artifact": artifact(source),
                "artifact_inventory": artifact(inventory_path),
                "raw_bytes": sum(p["bytes"] for p in inventory),
                "host_snapshots": len(hosts),
                "dirty_host_snapshots": dirty_hosts,
                "measurement": raw,
            })
        assert len(identities) == 1, (group, identities)
        assert len(blocks) >= 5, (group, len(blocks))
        comparisons[group] = {
            "blocks": len(blocks),
            "median_after_before": {
                key: statistics.median(block["after_before"][key] for block in blocks)
                for key in blocks[0]["after_before"]
            },
            "block_after_before": [block["after_before"] for block in blocks],
            "minimum_phase_ns": min(
                span for block in blocks for leg in block["legs"] for span in leg["spans_ns"].values()
            ),
        }
    result = {
        "selection": "Every complete guarded block from every c0 attempt is retained, regardless of its value. Incomplete blocks remain in each raw inventory. Earlier production candidates are separate records.",
        "comparisons": comparisons,
        "runs": runs,
    }
    save(args.out / "measurements.json", result)
    print(json.dumps(comparisons, indent=2))


if __name__ == "__main__":
    main()
