#!/usr/bin/env python3
"""Remove owned build scratch only after retained baseline replay passes."""
import datetime
import json
from pathlib import Path
import shlex
import shutil

import capture

PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]

def read(name):
    return json.loads((PACKET / name).read_text())

def check_file(record):
    actual = capture.descriptor(Path(record["path"]))
    for field in ("sha256", "bytes"):
        if actual[field] != record[field]:
            raise ValueError(f"retained {field} mismatch: {record['path']}")

def main():
    build = read("build.json")
    if any(read(name)["status"] != "pass" for name in
           ("build.json", "verification.json", "capture-review.json", "source-archive-replay.json")):
        raise ValueError("build and replay must pass before cleanup")
    capture.load_plan(PACKET / "plan.json")
    source = Path("/tmp/litchi-0885-source")
    target = ROOT / "target/litchi-0885-build"
    if Path(build["source_root"]) != source or Path(build["target_root"]) != target:
        raise ValueError("unexpected owned build paths")
    for path in (source, target):
        if path.is_symlink() or (path / ".owner-0885").read_text() != "performance baseline 0885\n":
            raise ValueError(f"invalid cleanup ownership marker: {path}")
    check_file(build["source_manifest"])
    check_file(build["source_archive"])
    manifest = read("source.json")
    dependency_counts = {}
    for mode, binary in build["binaries"].items():
        check_file(binary)
        check_file(binary["depfile"])
        check_file(binary["build_receipt"])
        receipt = json.loads(Path(binary["build_receipt"]["path"]).read_text())
        check_file(receipt["log"])
        line = Path(binary["depfile"]["path"]).read_text().splitlines()[0]
        dependencies = shlex.split(line.split(": ", 1)[1])
        for name in dependencies:
            relative = Path(name).resolve().relative_to(source).as_posix()
            if relative not in manifest["files"]:
                raise ValueError(f"unmanifested source dependency: {name}")
        dependency_counts[mode] = len(dependencies)
    removed = []
    for path in (source, target):
        shutil.rmtree(path)
        if path.exists():
            raise ValueError(f"cleanup did not remove {path}")
        removed.append(str(path))
    capture._write_json(PACKET / "cleanup.json", {
        "status": "pass", "finished_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "removed_owned_scratch": removed, "source_dependency_counts_verified": dependency_counts,
        "retained_local_artifacts": [build["source_archive"]["path"], *[row["path"] for row in build["binaries"].values()]],
        "registered_worktrees_created": 0,
        "scope": "Only owned0885 source/build directories removed; historical recovery and other teams' worktrees preserved.",
    })
    print("Owned source/build scratch removed; captured artifacts retained.")

if __name__ == "__main__":
    main()
