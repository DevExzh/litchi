#!/usr/bin/env python3
"""Bind generated file inputs, actual filesystems, and matched build sources.

Run through gate.py before formal captures. This supplements the frozen
workload protocol with observed execution inputs; it does not replace it.
"""

import argparse
import json
from pathlib import Path

import measure_routes as routes
from common import ROOT, REPO, TEMP, meta, now, read, sha, snapshot, write
from record_environment import command


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args()
    attempt = routes.base._attempt(args.attempt)
    destination = ROOT / "route-attempts" / attempt / "execution-inputs.json"
    routes.require(not destination.exists(), f"refusing to replace {destination}")
    _, protocol_hash = routes.load_protocol()
    builds = routes._load_builds(attempt, protocol_hash)
    current_source = snapshot()
    routes.require(current_source == builds["normal"]["source_after"], "execution sources differ from matched build inputs")
    prep_path = ROOT / "axis-input-origin" / attempt / "receipt.json"
    prep = read(prep_path)
    routes.require(prep["status"] == "pass" and prep["attempt"] == attempt, "input preparation did not pass")
    routes.require(prep["protocol_sha256"] == protocol_hash and prep["binary"] == builds["normal"]["binary"], "input preparation build/protocol differs")
    routes.require(prep["driver_sha256"] == sha(ROOT / "prepare_axis_inputs.py"), "preparation helper changed")
    for export in prep["exports"]:
        for name, expected in export["artifacts"].items():
            routes.require(meta(prep_path.parent / name) == expected, f"input export changed: {name}")
    inputs = []
    expected_paths = {arm["input_file"] for arm in routes.AXIS_ARMS if arm["input_mode"] == "file"}
    routes.require(len(prep["staged"]) == len(expected_paths) and {item["path"] for item in prep["staged"]} == expected_paths, "prepared file-input inventory differs")
    for item in prep["staged"]:
        path = ROOT / item["path"]
        routes.require(path.is_file() and not path.is_symlink(), f"input must remain a regular file: {path}")
        expected = {key: item[key] for key in ("bytes", "sha256")}
        routes.require(meta(path) == expected and meta(ROOT / item["origin"]) == expected, f"prepared file-input bytes changed: {path}")
        info = path.stat()
        inputs.append({**item, "absolute_path": str(path.resolve()), "device": info.st_dev, "inode": info.st_ino})
    filesystems = {}
    for name, path in (("workspace", REPO), ("evidence", ROOT), ("scratch", TEMP)):
        observation = command(["findmnt", "--json", "--target", str(path), "--output", "TARGET,SOURCE,FSTYPE,OPTIONS"])
        routes.require(observation["status"] == "pass", f"{name} mount observation unavailable")
        filesystems[name] = {"path": str(path.resolve()), "device": path.stat().st_dev, "mount": observation}
    value = {
        "schema": "docx-route-execution-inputs-v1", "attempt": attempt, "recorded_utc": now(),
        "driver_sha256": sha(Path(__file__)),
        "protocol": {"path": "route-protocol.json", **meta(routes.protocol_path())},
        "machine": {"path": "machine.json", **meta(ROOT / "machine.json")},
        "source": current_source,
        "builds": {role: {"path": f"route-attempts/{attempt}/build-{role}.json", **meta(ROOT / "route-attempts" / attempt / f"build-{role}.json")} for role in routes.ROLES},
        "preparation": {"path": prep_path.relative_to(ROOT).as_posix(), **meta(prep_path)},
        "file_inputs": inputs, "filesystems": filesystems,
        "replay_directory_roots": [f"route-pilots/{attempt}", f"route-captures/{attempt}"],
        "replay_filesystem": "evidence", "cold_cache_claim": False,
        "source_custody_scope": "matched build inputs and execution snapshot; the workload protocol has no separate Rust-source freeze field",
    }
    write(destination, value)
    print(json.dumps({"output": str(destination), "source": current_source["sha256"], "file_inputs": len(inputs)}), flush=True)


if __name__ == "__main__":
    main()
