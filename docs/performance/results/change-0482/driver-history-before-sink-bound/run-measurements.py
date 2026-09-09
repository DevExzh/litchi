#!/usr/bin/env python3
"""Run the frozen same-build XML reader comparison.

The executable is deliberately supplied by the caller.  This driver never
builds it and never reuses a capture path: every process gets an exclusive
started record, stdout/stderr files, GNU-time resource file, report path, and
terminal receipt.  A failed process is retained and the remaining group is
still attempted so that the evidence says exactly what was tried.
"""

from __future__ import annotations

import argparse
import fcntl
import hashlib
import json
import os
from pathlib import Path
import subprocess
from typing import Any

from common import ENV, ENV_KEYS, ROOT, REPO, TEMP, meta, now, read, sha, write


SCHEMA = "xml-stream-audit-comparison-v1"
REPORT_SCHEMA = "litchi.xml-stream-audit.v1"
SIZES = (64 * 1024, 8 * 1024 * 1024, 128 * 1024 * 1024)
MODES = ("materialized", "streaming")
INSTRUMENTATIONS = ("normal", "allocator")
ARMS = ("materialized", "streaming")
REPEATS = (1, 2)
SAMPLES = 30
WARMUPS = 3
CPU = 2
DEFAULT_REQUIRED_VALIDATION_LABELS = ("build-normal", "build-allocator")
XML_ASSET_SCHEMA = "xml-stream-audit-xml-assets-v1"
XML_ASSET_INVENTORY_PATH = "xml-assets.json"
XML_ASSET_COUNT = 77
FUZZ_SCHEMA = "xml-stream-audit-fuzz-evidence-v1"
FUZZ_SEED_COUNT = 11
FUZZ_DATA_PREFIX = "fuzz/final"


def _exclusive_write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    write(path, value)


def _snapshot() -> dict[str, Any]:
    # Importing the gate keeps source custody identical to the validation
    # receipts and avoids retaining a second copy of the repository manifest.
    from gate import snapshot

    return snapshot()


def _binary(path: Path) -> dict[str, Any]:
    if not path.is_file() or not os.access(path, os.X_OK):
        raise SystemExit(f"binary is not an executable regular file: {path}")
    return dict(path=str(path.resolve()), **meta(path))


def _binary_after(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        return dict(path=str(path.resolve()), **meta(path))
    except OSError:
        return None


def _script_digests() -> dict[str, str]:
    names = (
        "common.py",
        "gate.py",
        "build.py",
        "run-gates.py",
        "run-measurements.py",
        "analyze.py",
        "verify.py",
        "test_evidence.py",
        "fuzz.py",
    )
    return {name: sha(ROOT / name) for name in names if (ROOT / name).is_file()}


def _inventory(directory: Path) -> dict[str, dict[str, Any]]:
    if not directory.is_dir():
        raise SystemExit(f"missing retained fuzz directory: {directory}")
    values: dict[str, dict[str, Any]] = {}
    for path in sorted(directory.rglob("*")):
        if path.is_symlink():
            raise SystemExit(f"fuzz directory contains a symlink: {path}")
        if path.is_file():
            values[path.relative_to(directory).as_posix()] = dict(meta(path))
    return values


def _inventory_digest(values: dict[str, dict[str, Any]]) -> str:
    encoded = json.dumps(values, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def _fuzz_file_reference(relative: str) -> dict[str, Any]:
    path = ROOT / relative
    if not path.is_file():
        raise SystemExit(f"missing retained fuzz file: {path}")
    return dict(path=relative, **meta(path))


def _fuzz_inventory_reference(relative: str, minimum_files: int) -> dict[str, Any]:
    values = _inventory(ROOT / relative)
    if len(values) < minimum_files:
        raise SystemExit(f"{relative}: retained fuzz inventory has too few files")
    return {"path": relative, "files": len(values), "sha256": _inventory_digest(values)}


def _fuzz_reference() -> dict[str, Any]:
    """Bind freeze to the completed retained ASan/libFuzzer evidence."""

    return {
        "schema": FUZZ_SCHEMA,
        "cwd": str(REPO.resolve()),
        "helper": _fuzz_file_reference("fuzz.py"),
        "prepared": _fuzz_file_reference(f"{FUZZ_DATA_PREFIX}/prepared.json"),
        "build": _fuzz_file_reference(f"{FUZZ_DATA_PREFIX}/build.json"),
        "smoke": _fuzz_file_reference(f"{FUZZ_DATA_PREFIX}/smoke.json"),
        "manifest": _fuzz_file_reference(f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.toml"),
        "lock": _fuzz_file_reference(f"{FUZZ_DATA_PREFIX}/build-inputs/Cargo.lock.txt"),
        "seed_inventory": _fuzz_inventory_reference("fuzz/seeds", FUZZ_SEED_COUNT),
        "post_run_inventory": _fuzz_inventory_reference(f"{FUZZ_DATA_PREFIX}/post-run", 1),
    }


def _xml_asset_reference() -> dict[str, Any]:
    """Bind the frozen protocol to every XML source asset compiled into the binary."""

    path = ROOT / XML_ASSET_INVENTORY_PATH
    if not path.is_file():
        raise SystemExit(f"missing XML asset inventory: {path}")
    inventory = read(path)
    assets = inventory.get("assets") if isinstance(inventory, dict) else None
    if (
        not isinstance(inventory, dict)
        or set(inventory) != {"schema", "files", "assets"}
        or inventory.get("schema") != XML_ASSET_SCHEMA
        or inventory.get("files") != XML_ASSET_COUNT
        or not isinstance(assets, dict)
        or len(assets) != XML_ASSET_COUNT
    ):
        raise SystemExit("xml-assets.json: inventory schema or count differs")
    current: dict[str, str] = {}
    for candidate in (REPO / "crates").rglob("*.xml"):
        relative = candidate.relative_to(REPO)
        if candidate.is_file() and "src" in relative.parts:
            current[relative.as_posix()] = sha(candidate)
    if assets != current:
        raise SystemExit("xml-assets.json: inventory does not match current XML source assets")
    return {"path": XML_ASSET_INVENTORY_PATH, "sha256": sha(path), "files": XML_ASSET_COUNT}


def _build_references(binaries: dict[str, dict[str, Any]]) -> dict[str, dict[str, str]]:
    references: dict[str, dict[str, str]] = {}
    for instrumentation in INSTRUMENTATIONS:
        path = ROOT / "builds" / f"{instrumentation}.json"
        if not path.is_file():
            raise SystemExit(f"missing successful {instrumentation} build receipt: {path}")
        build = read(path)
        if not isinstance(build, dict) or build.get("schema") != "xml-stream-audit-build-v1":
            raise SystemExit(f"{path}: build receipt schema differs")
        if build.get("instrumentation") != instrumentation:
            raise SystemExit(f"{path}: build instrumentation differs")
        if build.get("source_snapshot") != _snapshot():
            raise SystemExit(f"{path}: build source snapshot differs from freeze source")
        if build.get("binary") != binaries[instrumentation]:
            raise SystemExit(f"{path}: build binary differs from explicit freeze binary")
        references[instrumentation] = {"path": path.relative_to(ROOT).as_posix(), "sha256": sha(path)}
    return references


def _capture_rows(instrumentation: str, binaries: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    """Return the exact A1/B1/B2/A2 process order for one build."""

    if instrumentation not in INSTRUMENTATIONS:
        raise ValueError(f"unknown instrumentation {instrumentation!r}")
    forward = [(mode, size) for mode in MODES for size in SIZES]
    rows: list[dict[str, Any]] = []
    sequences = (
        ("materialized", 1, [("materialized", size) for size in SIZES]),
        ("streaming", 1, [("streaming", size) for size in SIZES]),
        ("streaming", 2, [("streaming", size) for size in reversed(SIZES)]),
        ("materialized", 2, [("materialized", size) for size in reversed(SIZES)]),
    )
    for arm, repeat, sequence in sequences:
        for mode, size in sequence:
            label = f"{instrumentation}-{arm}-r{repeat}-{size}"
            prefix = ROOT / "captures" / label
            binary = binaries[instrumentation]["path"]
            rows.append(
                dict(
                    label=label,
                    instrumentation=instrumentation,
                    arm=arm,
                    mode=mode,
                    size_bytes=size,
                    repeat=repeat,
                    argv=[
                        "/usr/bin/time",
                        "-v",
                        "-o",
                        str(prefix.with_suffix(".resource").resolve()),
                        "/usr/bin/taskset",
                        "-c",
                        str(CPU),
                        binary,
                        "--mode",
                        mode,
                        "--sizes",
                        str(size),
                        "--samples",
                        str(SAMPLES),
                        "--warmup",
                        str(WARMUPS),
                        "--json",
                        str(prefix.with_suffix(".report.json").resolve()),
                    ],
                )
            )
    return rows


def _pilot_rows(instrumentation: str, binary: dict[str, Any], attempt: int) -> list[dict[str, Any]]:
    if attempt <= 0:
        raise ValueError("pilot attempt must be positive")
    rows: list[dict[str, Any]] = []
    for mode in MODES:
        for size in SIZES:
            label = f"pilot-{instrumentation}-a{attempt}-{mode}-{size}"
            prefix = ROOT / "pilots" / label
            rows.append(
                dict(
                    label=label,
                    instrumentation=instrumentation,
                    arm=mode,
                    mode=mode,
                    size_bytes=size,
                    repeat=0,
                    samples=1,
                    warmups=1,
                    argv=[
                        "/usr/bin/time",
                        "-v",
                        "-o",
                        str(prefix.with_suffix(".resource").resolve()),
                        "/usr/bin/taskset",
                        "-c",
                        str(CPU),
                        binary["path"],
                        "--mode",
                        mode,
                        "--sizes",
                        str(size),
                        "--samples",
                        "1",
                        "--warmup",
                        "1",
                        "--json",
                        str(prefix.with_suffix(".report.json").resolve()),
                    ],
                )
            )
    return rows


def _assert_new(prefix: Path) -> None:
    suffixes = (".started.json", ".stdout", ".stderr", ".resource", ".report.json", ".json")
    existing = [str(prefix.with_suffix(s)) for s in suffixes if prefix.with_suffix(s).exists()]
    if existing:
        raise SystemExit("refusing to overwrite retained measurement files: " + ", ".join(existing))


def _artifact_record(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    return dict(path=str(path.resolve()), **meta(path))


def _run_one(row: dict[str, Any], binary: dict[str, Any], protocol_hash: str | None) -> int:
    prefix = ROOT / ("pilots" if row["label"].startswith("pilot-") else "captures") / row["label"]
    _assert_new(prefix)
    prefix.parent.mkdir(parents=True, exist_ok=True)
    binary_path = Path(binary["path"])
    source_before = _snapshot()
    binary_before = _binary(binary_path)
    started = dict(
        schema=SCHEMA,
        capture=row,
        cwd=str(REPO),
        started_utc=now(),
        protocol_sha256=protocol_hash,
        source_before=source_before,
        binary_before=binary_before,
        environment={key: ENV[key] for key in ENV_KEYS},
    )
    _exclusive_write(prefix.with_suffix(".started.json"), started)
    stdout = prefix.with_suffix(".stdout")
    stderr = prefix.with_suffix(".stderr")
    with stdout.open("xb") as out, stderr.open("xb") as err:
        try:
            result = subprocess.run(row["argv"], cwd=REPO, env=ENV, stdout=out, stderr=err)
            exit_code = result.returncode
        except OSError as error:
            err.write(f"measurement process could not be started: {error}\n".encode())
            exit_code = 127
    source_after = _snapshot()
    terminal = dict(
        started,
        exit_code=exit_code,
        finished_utc=now(),
        source_after=source_after,
        binary_after=_binary_after(binary_path),
        source_unchanged=source_before == source_after,
        binary_unchanged=_binary_after(binary_path) == binary_before,
        artifacts={
            name: value
            for name, value in (
                (stdout.name, _artifact_record(stdout)),
                (stderr.name, _artifact_record(stderr)),
                (prefix.with_suffix(".resource").name, _artifact_record(prefix.with_suffix(".resource"))),
                (prefix.with_suffix(".report.json").name, _artifact_record(prefix.with_suffix(".report.json"))),
            )
            if value is not None
        },
    )
    _exclusive_write(prefix.with_suffix(".json"), terminal)
    print(f"{row['label']} {exit_code} source unchanged {terminal['source_unchanged']}", flush=True)
    return exit_code


def _freeze(args: argparse.Namespace) -> None:
    if not args.normal_binary or not args.allocator_binary:
        raise SystemExit("--freeze requires --normal-binary and --allocator-binary")
    binaries = {
        "normal": _binary(Path(args.normal_binary).resolve()),
        "allocator": _binary(Path(args.allocator_binary).resolve()),
    }
    captures = []
    for instrumentation in INSTRUMENTATIONS:
        captures.extend(_capture_rows(instrumentation, binaries))
    provided_labels = list(args.required_validation_label or ())
    if len(set(provided_labels)) != len(provided_labels) or not all(provided_labels):
        raise SystemExit("--required-validation-label values must be non-empty and unique")
    required_labels = list(DEFAULT_REQUIRED_VALIDATION_LABELS)
    required_labels.extend(label for label in provided_labels if label not in required_labels)
    protocol = dict(
        schema=SCHEMA,
        report_schema=REPORT_SCHEMA,
        version=1,
        samples=SAMPLES,
        warmups=WARMUPS,
        cpu=CPU,
        sizes=list(SIZES),
        modes=list(MODES),
        arms=list(ARMS),
        repeats=list(REPEATS),
        sequence="A1/B1/B2/A2",
        comparison="same executable and source revision; materialized versus streaming route only",
        performance_claim="primitive-enabler-only",
        no_cross_revision_claim=True,
        process_rss="GNU time -v Maximum resident set size, per process; setup and teardown included",
        normal_and_allocator_timings_separate=True,
        pair_review_percent=5,
        required_validation_labels=required_labels,
        environment={key: ENV[key] for key in ENV_KEYS},
        source_snapshot=_snapshot(),
        binaries=binaries,
        builds=_build_references(binaries),
        scripts=_script_digests(),
        xml_asset_inventory=_xml_asset_reference(),
        fuzz=_fuzz_reference(),
        captures=captures,
    )
    _exclusive_write(ROOT / "protocol.json", protocol)
    print(f"Frozen {len(captures)} formal captures ({len(captures) * SAMPLES} measured samples).")


def _run_formal(args: argparse.Namespace) -> int:
    protocol_path = ROOT / "protocol.json"
    protocol = read(protocol_path)
    if protocol.get("schema") != SCHEMA:
        raise SystemExit("protocol schema differs")
    if protocol.get("scripts", {}).get("run-measurements.py") != sha(Path(__file__)):
        raise SystemExit("run-measurements.py changed after protocol freeze")
    instrumentation = args.instrumentation
    binary_spec = protocol["binaries"][instrumentation]
    binary_path = Path(args.binary or binary_spec["path"]).resolve()
    current_binary = _binary(binary_path)
    if current_binary != binary_spec:
        raise SystemExit(f"{instrumentation} binary differs from frozen protocol")
    rows = [row for row in protocol["captures"] if row["instrumentation"] == instrumentation]
    if args.arm:
        rows = [row for row in rows if row["arm"] == args.arm]
    if args.repeat:
        rows = [row for row in rows if row["repeat"] == args.repeat]
    if not rows:
        raise SystemExit("no formal captures selected")
    expected = _capture_rows(instrumentation, protocol["binaries"])
    if rows != [row for row in expected if row["label"] in {item["label"] for item in rows}]:
        raise SystemExit("protocol capture rows do not match the frozen runner plan")
    protocol_hash = sha(protocol_path)
    failures = 0
    # The caller may invoke one arm/repeat per process.  We still retain every
    # attempted process in the selected group before returning failure.
    for row in rows:
        row = dict(row)
        failures |= int(_run_one(row, binary_spec, protocol_hash) != 0)
    return failures


def _run_pilot(args: argparse.Namespace) -> int:
    path = Path(args.binary).resolve() if args.binary else None
    if path is None:
        raise SystemExit("--pilot requires --binary")
    binary = _binary(path)
    rows = _pilot_rows(args.instrumentation, binary, args.attempt)
    failures = 0
    for row in rows:
        failures |= int(_run_one(row, binary, None) != 0)
    return failures


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--freeze", action="store_true")
    parser.add_argument("--pilot", action="store_true")
    parser.add_argument("--instrumentation", choices=INSTRUMENTATIONS)
    parser.add_argument("--binary")
    parser.add_argument("--normal-binary")
    parser.add_argument("--allocator-binary")
    parser.add_argument("--arm", choices=ARMS)
    parser.add_argument("--repeat", type=int, choices=REPEATS)
    parser.add_argument("--attempt", type=int, default=1)
    parser.add_argument(
        "--required-validation-label",
        action="append",
        help="validation gate label to require in the frozen bundle (repeat for each; both build labels are mandatory)",
    )
    return parser


def main() -> int:
    args = _parser().parse_args()
    if args.freeze:
        if any(value is not None for value in (args.instrumentation, args.binary, args.arm, args.repeat)):
            raise SystemExit("--freeze cannot select a run")
        _freeze(args)
        return 0
    if args.pilot:
        if args.instrumentation is None or args.arm or args.repeat:
            raise SystemExit("--pilot requires --instrumentation and cannot select arm/repeat")
        TEMP.mkdir(parents=True, exist_ok=True)
        with (TEMP / "cpu.lock").open("a") as lock:
            fcntl.flock(lock, fcntl.LOCK_EX)
            return _run_pilot(args)
    if args.instrumentation is None:
        raise SystemExit("formal run requires --instrumentation")
    if args.arm is None or args.repeat is None:
        raise SystemExit("formal run requires --arm and --repeat")
    TEMP.mkdir(parents=True, exist_ok=True)
    with (TEMP / "cpu.lock").open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        return _run_formal(args)


if __name__ == "__main__":
    raise SystemExit(main())
