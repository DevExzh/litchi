#!/usr/bin/env python3
"""Matched six-row normal-process RSS captures; invoke under cpu.lock."""

import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
FLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
ROLE_FOR_LANE = {
    "rss-A1": "control",
    "rss-A2": "control",
    "rss-B1": "candidate",
    "rss-B2": "candidate",
}


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    lane = sys.argv[1]
    if lane not in ROLE_FOR_LANE:
        raise SystemExit(f"unsupported RSS lane: {lane}")
    role = ROLE_FOR_LANE[lane]
    protocol_path = ROOT / "rss-protocol.json"
    protocol = json.loads(protocol_path.read_text(encoding="utf-8"))
    if sha(Path(__file__)) != protocol["capture_driver_sha256"]:
        raise SystemExit("RSS capture driver hash differs from frozen protocol")
    main_protocol = ROOT / "protocol.json"
    if sha(main_protocol) != protocol["main_protocol_sha256"]:
        raise SystemExit("main protocol hash differs from frozen RSS protocol")

    binding = json.loads((ROOT / f"{role}-binding.json").read_text(encoding="utf-8"))
    binary = Path(binding["binary_path"])
    worktree = Path(binding["build_path"])
    if subprocess.check_output(["git", "status", "--porcelain"], cwd=worktree).strip():
        raise SystemExit("RSS capture source tree is dirty before checkout")
    subprocess.run(["git", "checkout", "--detach", binding["revision"]], cwd=worktree, check=True)

    def check_clean_role() -> None:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=worktree, text=True).strip()
        if revision != binding["revision"]:
            raise SystemExit("RSS capture revision differs from role binding")
        if subprocess.check_output(["git", "status", "--porcelain"], cwd=worktree).strip():
            raise SystemExit("RSS capture source tree is dirty")

    check_clean_role()
    if sha(binary) != binding["binary_sha256"]:
        raise SystemExit("RSS capture binary hash differs from role binding")
    manifest = ROOT / binding["source_manifest"]
    if sha(manifest) != binding["source_manifest_sha256"]:
        raise SystemExit("RSS capture source manifest hash differs")
    for name, digest in json.loads(manifest.read_text(encoding="utf-8")).items():
        if sha(worktree / name) != digest:
            raise SystemExit(f"RSS capture source digest differs: {name}")
    for name, digest in binding["included_fixtures"].items():
        if sha(worktree / name) != digest:
            raise SystemExit(f"RSS capture fixture digest differs: {name}")

    output = ROOT / lane
    output.mkdir(exist_ok=False)
    samples = int(protocol["samples"])
    warmups = int(protocol["warmups"])
    cases = ",".join(protocol["cases"])
    shapes = ",".join(protocol["shapes"])
    workload = [
        str(binary),
        "--workers",
        str(protocol["workers"]),
        "--warmup",
        str(warmups),
        "--samples",
        str(samples),
        "--json",
        str(output / "report.json"),
        "--corpus-manifest",
        str(output / "corpus-catalog.json"),
        "--case",
        cases,
        "--xlsx-shape",
        shapes,
    ]
    argv = [
        "taskset",
        "-c",
        str(protocol["cpu"]),
        "/usr/bin/time",
        "-v",
        "-o",
        str(output / "resource.log"),
        *workload,
    ]
    env = dict(
        os.environ,
        RUSTUP_TOOLCHAIN="1.98.1",
        RUSTFLAGS=FLAGS,
        CARGO_PROFILE_RELEASE_DEBUG="1",
        CARGO_INCREMENTAL="0",
        CARGO_BUILD_JOBS="4",
        DEBUGINFOD_URLS="",
        LC_ALL="C",
        PYTHONDONTWRITEBYTECODE="1",
    )
    receipt = {
        "schema": "litchi-0470-rss-capture-v1",
        "lane": lane,
        "role": role,
        "revision": binding["revision"],
        "binary_sha256": sha(binary),
        "binding_sha256": sha(ROOT / f"{role}-binding.json"),
        "driver_sha256": sha(Path(__file__)),
        "protocol_sha256": sha(protocol_path),
        "main_protocol_sha256": sha(main_protocol),
        "argv": argv,
        "cwd": str(worktree),
        "samples": samples,
        "warmups": warmups,
        "environment": {
            key: env[key]
            for key in (
                "RUSTUP_TOOLCHAIN",
                "RUSTFLAGS",
                "CARGO_PROFILE_RELEASE_DEBUG",
                "DEBUGINFOD_URLS",
                "LC_ALL",
            )
        },
        "started_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "clean_before": True,
    }
    (output / "started.json").write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
    with (output / "stdout.log").open("w", encoding="utf-8") as stdout, (output / "stderr.log").open("w", encoding="utf-8") as stderr:
        result = subprocess.run(argv, cwd=worktree, env=env, stdout=stdout, stderr=stderr)
    check_clean_role()
    if sha(binary) != binding["binary_sha256"]:
        raise SystemExit("RSS capture binary changed")
    report_ok = False
    if result.returncode == 0:
        report = json.loads((output / "report.json").read_text(encoding="utf-8"))
        report_ok = (
            report["environment"]["git_revision"] == binding["revision"]
            and report["environment"]["git_worktree_dirty"] is False
        )
    receipt.update(
        exit_code=result.returncode,
        clean_after=True,
        binary_unchanged=True,
        report_metadata_matches_clean_role=report_ok,
        finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        artifacts={
            path.name: {"sha256": sha(path), "bytes": path.stat().st_size}
            for path in output.iterdir()
            if path.is_file()
        },
    )
    (output / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"lane": lane, "exit_code": result.returncode}), flush=True)
    raise SystemExit(result.returncode if result.returncode else int(not report_ok))


if __name__ == "__main__":
    main()
