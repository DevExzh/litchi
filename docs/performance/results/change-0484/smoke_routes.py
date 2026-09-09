#!/usr/bin/env python3
"""Diagnostic route executions; one sample is correctness evidence, not timing evidence.

Run through gate.py to retain source snapshots and serialize heavy commands.
This command neither freezes a protocol nor builds a binary. Failed files are
retained; only verified empty replay directories are removed.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import subprocess

import measure_routes as routes
from common import ENV, REPO, ROOT, TEMP, meta, now, write


def run(binary_path: Path, role: str, attempt: str) -> int:
    attempt = routes.base._attempt(attempt)
    binary_path = binary_path.resolve(strict=True)
    binary = {"path": str(binary_path), **meta(binary_path)}
    destination = ROOT / "route-diagnostics" / attempt
    destination.mkdir(parents=True, exist_ok=False)
    scratch = TEMP / "route-diagnostics" / attempt
    scratch.mkdir(parents=True, exist_ok=False)
    bindings = {
        "binary": binary,
        "role": role,
        "attempt": attempt,
        "diagnostic_only": True,
        "samples": 1,
        "warmups": 1,
        "driver": meta(Path(__file__)),
        "validators": routes._script_hashes(),
        "machine": routes._machine_binding(required=True),
    }
    write(destination / "started.json", dict(bindings, started_utc=now()))
    failures = []
    receipts = []
    for case_label in ("s64-a64-short-c64", "s64-a64-near-c64"):
        case = dict(routes.ROUTE_CASE_BY_LABEL[case_label])
        for spec in routes.ROUTES:
            label = f"{spec.name}-{case_label}"
            directory = destination / label
            directory.mkdir()
            replay = scratch / label if spec.name == "file_store" else None
            if replay is not None:
                replay.mkdir()
            report = directory / "report.json"
            resource = directory / "resource.txt"
            stdout = directory / "stdout.txt"
            stderr = directory / "stderr.txt"
            argv = routes._route_argv(
                binary, case, spec, samples=1, warmups=1, report=report,
                resource=resource, replay_dir=replay,
            )
            started = dict(
                label=label, argv=argv, cwd=str(REPO), started_utc=now(),
                case=case, route=spec.name,
            )
            write(directory / "started.json", started)
            error = None
            exit_code = None
            try:
                with stdout.open("xb") as out, stderr.open("xb") as err:
                    process = subprocess.run(
                        argv, cwd=REPO, env=ENV, stdout=out, stderr=err,
                        check=False,
                    )
                exit_code = process.returncode
                if exit_code:
                    raise RuntimeError(f"child exited {exit_code}")
                routes.check_route_report(
                    report, role, case, spec, samples=1, warmups=1,
                    binary=binary, argv=argv, replay_dir=replay,
                )
                if replay is not None:
                    if any(replay.iterdir()):
                        raise RuntimeError("successful route retained replay files")
                    replay.rmdir()
            except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as caught:
                error = f"{type(caught).__name__}: {caught}"
                failures.append({"label": label, "error": error})
            receipt = dict(
                started, finished_utc=now(), exit_code=exit_code,
                passed=error is None, error=error,
                artifacts={p.name: meta(p) for p in (report, resource, stdout, stderr) if p.is_file()},
            )
            write(directory / "receipt.json", receipt)
            receipts.append({"label": label, **meta(directory / "receipt.json")})
            print(f"{label}: {'PASS' if error is None else error}", flush=True)
    unchanged = meta(binary_path) == {key: binary[key] for key in ("bytes", "sha256")}
    if not unchanged:
        failures.append({"label": "binary", "error": "executable changed during diagnostics"})
    empty = not any(scratch.iterdir())
    if empty:
        scratch.rmdir()
    else:
        failures.append({"label": "scratch", "error": "diagnostic scratch is not empty; files retained"})
    write(destination / "result.json", dict(
        bindings, finished_utc=now(), passed=not failures, failures=failures,
        receipts=receipts, binary_unchanged=unchanged, scratch_removed=empty,
    ))
    return int(bool(failures))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--role", choices=routes.ROLES, required=True)
    parser.add_argument("--attempt", required=True)
    args = parser.parse_args()
    raise SystemExit(run(args.binary, args.role, args.attempt))


if __name__ == "__main__":
    main()
