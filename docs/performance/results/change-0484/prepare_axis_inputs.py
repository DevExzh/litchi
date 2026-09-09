#!/usr/bin/env python3
"""Export and authenticate the frozen axis file-input corpus outside captures.

Run through gate.py to serialize this setup with builds and measurements.
Exports and receipts are retained; no setup sample is a formal measurement.
"""

import argparse
import hashlib
from pathlib import Path
import shutil
import subprocess
import zipfile

import measure_routes as routes
from common import ENV, REPO, ROOT, meta, now, read, sha, write


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args()
    attempt = routes.base._attempt(args.attempt)
    _, protocol_hash = routes.load_protocol()
    builds = routes._load_builds(attempt, protocol_hash)
    binary = builds["normal"]["binary"]
    directory = ROOT / "axis-input-origin" / attempt
    directory.mkdir(parents=True, exist_ok=False)
    started = {
        "schema": "docx-axis-input-preparation-v1", "status": "running",
        "started_utc": now(), "attempt": attempt,
        "driver_sha256": sha(Path(__file__)),
        "protocol_sha256": protocol_hash, "binary": binary,
        "build": {"path": f"route-attempts/{attempt}/build-normal.json",
                  "sha256": sha(ROOT / "route-attempts" / attempt / "build-normal.json")},
        "scope": "excluded corpus setup; no formal timing samples",
    }
    write(directory / "started.json", started)
    sources = {}
    exports = []
    staged = []
    try:
        for count in sorted({routes.CASE_BY_LABEL[label]["source_count"] for label in routes.AXIS_WORKLOAD_LABELS}):
            child = directory / f"source-{count}"
            child.mkdir()
            argv = ["/usr/bin/taskset", "-c", str(routes.base.CPU), binary["path"],
                    "--source-counts", str(count), "--authored-counts", "64",
                    "--chunks", "64", "--text", "short", "--samples", "1", "--warmups", "1",
                    "--fixture-dir", str(child), "--json", str(child / "report.json")]
            write(child / "started.json", {"argv": argv, "cwd": str(REPO), "started_utc": now()})
            with (child / "stdout.txt").open("xb") as out, (child / "stderr.txt").open("xb") as err:
                process = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err, check=False)
            routes.require(process.returncode == 0, f"source export {count} exited {process.returncode}")
            manifests = list(child.glob("*-hashes.json"))
            routes.require(len(manifests) == 1, "fixture manifest cardinality differs")
            manifest = read(manifests[0])
            routes.require(manifest["source_count"] == count, "fixture source count differs")
            source = child / manifest["source"]["file"]
            routes.require(meta(source) == {key: manifest["source"][key] for key in ("bytes", "sha256")}, "fixture archive differs")
            with zipfile.ZipFile(source) as archive:
                xml = archive.read("word/document.xml")
            expected = routes.corpus_oracle.expected_case(count, 64, "fixed64", "short")["source"]
            routes.require(len(xml) == expected["main_xml_bytes"] and hashlib.sha256(xml).hexdigest() == expected["main_xml_sha256"], "independent source XML oracle differs")
            routes.require(manifest["source_main_xml_sha256"] == expected["main_xml_sha256"], "Rust fixture XML proof differs")
            sources[count] = source
            exports.append({"source_count": count, "argv": argv,
                            "artifacts": {p.relative_to(directory).as_posix(): meta(p) for p in child.iterdir() if p.is_file()}})
        for label in routes.AXIS_WORKLOAD_LABELS:
            arm = routes.AXIS_ARM_BY_LABEL[f"axis-input-file-{label}"]
            source = sources[routes.CASE_BY_LABEL[label]["source_count"]]
            target = ROOT / arm["input_file"]
            target.parent.mkdir(parents=True, exist_ok=True)
            with source.open("rb") as inp, target.open("xb") as out:
                shutil.copyfileobj(inp, out)
            routes.require(meta(target) == meta(source), "staged file differs from exported source")
            staged.append({"path": target.relative_to(ROOT).as_posix(), **meta(target),
                           "origin": source.relative_to(ROOT).as_posix()})
        write(directory / "receipt.json", dict(started, status="pass", finished_utc=now(), exports=exports, staged=staged))
        print(f"Prepared {len(staged)} authenticated file-input paths from {len(exports)} source exports", flush=True)
    except Exception as error:
        write(directory / "receipt.json", dict(started, status="failed", finished_utc=now(), error=f"{type(error).__name__}: {error}", exports=exports, staged=staged))
        raise


if __name__ == "__main__":
    main()
