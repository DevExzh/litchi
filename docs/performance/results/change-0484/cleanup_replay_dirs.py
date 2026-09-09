#!/usr/bin/env python3
"""Remove only verified empty replay directories from accepted formal lanes."""

import argparse
from pathlib import Path

import measure_routes as routes
from common import ROOT, meta, now, read, sha, write


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args()
    attempt = routes.base._attempt(args.attempt)
    output = ROOT / f"cleanup-replay-dirs-{attempt}.json"
    routes.require(not output.exists(), f"refusing to replace {output}")
    protocol, _ = routes.load_protocol()
    execution = read(ROOT / "route-attempts" / attempt / "execution-inputs.json")
    expected_device = execution["filesystems"]["evidence"]["device"]
    pending = []
    for directory, inventory in (("route-pilots", protocol["pilot_runs"]), ("route-captures", protocol["formal_runs"])):
        for run in inventory:
            if run["route"] != "file_store":
                continue
            receipt_path = ROOT / directory / attempt / run["label"] / "receipt.json"
            receipt = read(receipt_path)
            routes.require(receipt["status"] == "pass" and receipt["exit_code"] == 0, f"unaccepted child: {receipt_path}")
            routes.require(receipt["attempt"] == attempt and receipt["run"]["route"] == "file_store", "cleanup child identity differs")
            report_path = receipt_path.parent / "report.json"
            routes.require(meta(report_path) == receipt["artifacts"]["report.json"], f"report changed: {report_path}")
            report = read(report_path)
            routes.require(all(sample["replay"]["file_cleanup_verified"] is True for sample in report["cases"][0]["samples"]), "file cleanup was not verified")
            replay = receipt_path.parent / "replay"
            routes.require(replay.is_dir() and not replay.is_symlink(), f"replay directory missing or irregular: {replay}")
            routes.require(replay.resolve().is_relative_to(ROOT.resolve()), "replay directory escaped evidence root")
            routes.require(not any(replay.iterdir()), f"replay directory is not empty: {replay}")
            info = replay.stat()
            routes.require(info.st_dev == expected_device, "replay directory device differs from execution binding")
            pending.append((replay, {
                "path": replay.relative_to(ROOT).as_posix(),
                "receipt": {"path": receipt_path.relative_to(ROOT).as_posix(), **meta(receipt_path)},
                "device": info.st_dev, "inode": info.st_ino,
                "empty_before_removal": True, "removed": False,
            }))
    routes.require(len(pending) == 60, "expected exactly 60 accepted file-route directories")
    started = {"schema": "docx-replay-directory-cleanup-v1", "attempt": attempt,
               "started_utc": now(), "driver_sha256": sha(Path(__file__)),
               "scope": "empty replay directories only; generated corpus and diagnostic artifacts retained"}
    write(ROOT / f"cleanup-replay-dirs-{attempt}.started.json", dict(started, entries=[item for _, item in pending]))
    entries = []
    try:
        for path, item in pending:
            info = path.stat()
            routes.require(not path.is_symlink() and info.st_dev == item["device"] and info.st_ino == item["inode"], "replay directory identity changed before removal")
            path.rmdir()
            entries.append(dict(item, removed=True))
    except Exception as error:
        write(output, dict(started, status="failed", finished_utc=now(), error=str(error), entries=entries))
        raise
    write(output, dict(started, status="pass", finished_utc=now(), entries=entries))
    print(f"Removed {len(entries)} verified empty replay directories")


if __name__ == "__main__":
    main()
