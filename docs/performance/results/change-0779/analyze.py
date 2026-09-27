"""Offline replay and analysis for the 0779 XLSX ingress capture.

This module intentionally does not execute Cargo, the probe, a benchmark, or a
profiler.  It verifies the frozen packet, replays every retained receipt, and
derives statistics directly from the JSON reports and RSS gauges.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PREVIOUS = ROOT / "docs/performance/results/change-0778"
PHASES = ("open", "edit", "save", "lifecycle")
LEGS = ("before", "after")
METRICS = ("p50", "mean", "p95", "p99")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
HEX = frozenset("0123456789abcdefABCDEF")
PHYS_PKG = "crates/litchi-opc/src/phys_pkg.rs"
MARKER = "litchi-perf-0638-ordinary-save"
REPORT_SCHEMA = "litchi.xlsx.allocation-probe.v1"
REPORT_TOOL = "litchi-xlsx-allocation-attribution-probe-0779"


class ReplayError(RuntimeError):
    """Raised for any missing, stale, or contradictory evidence."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def _relocated(raw: Path) -> list[Path]:
    """Resolve absolute receipts retained from old worktree locations."""

    candidates: list[Path] = []
    parts = raw.parts
    for marker, base in (
        ("change-0779", PACKET),
        ("change-0778", PREVIOUS),
    ):
        if marker in parts:
            index = parts.index(marker)
            candidates.append(base.joinpath(*parts[index + 1:]))
    for marker in ("build-before", "build-after", "native", "allocation", "qualification"):
        if marker in parts:
            index = parts.index(marker)
            candidates.append(PACKET / marker / Path(*parts[index + 1:]))
    return candidates


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.extend(_relocated(raw))
        candidates.append(raw)
    else:
        text = value.replace("\\", "/")
        if text.startswith("docs/performance/results/change-0779/"):
            candidates.append(PACKET / text.split("change-0779/", 1)[1])
        elif text.startswith("docs/performance/results/change-0778/"):
            candidates.append(PREVIOUS / text.split("change-0778/", 1)[1])
        candidates.extend((PACKET / raw, PREVIOUS / raw))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            path = candidate.resolve()
            if packet_bound:
                try:
                    path.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return path
    path = (candidates[0] if candidates else raw).resolve(strict=False)
    if packet_bound:
        try:
            path.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return path


def artifact(
    value: Any,
    label: str,
    *,
    packet_bound: bool = True,
    allow_missing: bool = False,
) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi-0779-plan-v1", "plan schema changed")
    require(isinstance(plan.get("base"), str) and 7 <= len(plan["base"]) <= 64
            and all(c in HEX for c in plan["base"]), "plan base is invalid")
    require(plan.get("phases") == list(PHASES), "phase schedule changed")
    require(plan.get("review_percent") == 5, "review threshold changed")
    native = plan.get("native")
    allocation = plan.get("allocation")
    require(isinstance(native, dict) and isinstance(allocation, dict),
            "plan lane configuration is missing")
    require(native.get("blocks") == 6 and native.get("samples") == 30
            and native.get("warmup") == 3, "native lane configuration changed")
    require(allocation.get("blocks") == 2 and allocation.get("samples") == 3
            and allocation.get("warmup") == 0, "allocation lane configuration changed")
    orders = native.get("orders")
    require(isinstance(orders, list) and len(orders) == 6, "native order count changed")
    for index, order in enumerate(orders):
        require(order in (["before", "after"], ["after", "before"]),
                f"native block {index} order is invalid")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 2, "corpus count changed")
    ids: set[str] = set()
    for index, corpus in enumerate(corpora):
        require(isinstance(corpus, dict), f"corpus {index} is invalid")
        cid = corpus.get("id")
        require(isinstance(cid, str) and cid and cid not in ids, f"corpus {index} id is invalid")
        ids.add(cid)
        require(isinstance(corpus.get("source"), str), f"{cid} source is missing")
        positive_int(corpus.get("bytes"), f"{cid}.bytes")
        require(is_sha(corpus.get("sha256")), f"{cid}.sha256 is invalid")
        require(isinstance(corpus.get("sheet"), str) and corpus["sheet"],
                f"{cid}.sheet is missing")
    return plan


def load_small_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "small-control-plan.json")
    require(isinstance(plan, dict), "small-control plan is not an object")
    require(plan.get("schema") == "litchi-0779-small-controls-v1",
            "small-control plan schema changed")
    require(plan.get("phases") == ["open"], "small-control phase scope changed")
    require(plan.get("cpu") == 12 and plan.get("review_percent") == 5,
            "small-control host or review setting changed")
    native = plan.get("native")
    allocation = plan.get("allocation")
    require(isinstance(native, dict) and isinstance(allocation, dict),
            "small-control lane configuration is missing")
    require(native.get("blocks") == 6 and native.get("samples") == 30
            and native.get("warmup") == 3, "small native configuration changed")
    require(allocation.get("blocks") == 2 and allocation.get("samples") == 3
            and allocation.get("warmup") == 0, "small allocation configuration changed")
    orders = native.get("orders")
    require(isinstance(orders, list) and len(orders) == 6, "small order count changed")
    for order in orders:
        require(order in (["before", "after"], ["after", "before"]),
                "small-control order is invalid")
    items = plan.get("corpora")
    require(isinstance(items, list) and len(items) == 2,
            "small-control corpus count changed")
    seen: set[str] = set()
    for index, corpus in enumerate(items):
        require(isinstance(corpus, dict), f"small corpus {index} is invalid")
        cid = corpus.get("id")
        require(isinstance(cid, str) and cid not in seen and cid,
                f"small corpus {index} id is invalid")
        seen.add(cid)
        require(isinstance(corpus.get("source"), str) and corpus["source"],
                f"{cid} source is missing")
        positive_int(corpus.get("bytes"), f"{cid}.bytes")
        require(is_sha(corpus.get("sha256")), f"{cid}.sha256 is invalid")
        require(isinstance(corpus.get("sheet"), str) and corpus["sheet"],
                f"{cid}.sheet is missing")
        source = ROOT / corpus["source"]
        require(source.is_file() and not source.is_symlink(), f"{cid} fixture is missing")
        require(source.stat().st_size == corpus["bytes"] and sha256(source) == corpus["sha256"],
                f"{cid} fixture identity changed")
    return plan


def corpora(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["id"]: item for item in plan["corpora"]}


def _cleanup_has(cleanup: Any, receipt: dict[str, Any]) -> bool:
    expected = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
                receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(expected[0], str) and is_sha(expected[2])):
        return False
    if isinstance(cleanup, dict):
        if ((cleanup.get("path"), cleanup.get("bytes", cleanup.get("size")),
             cleanup.get("sha256", cleanup.get("digest"))) == expected):
            return True
        return any(_cleanup_has(value, receipt) for value in cleanup.values())
    if isinstance(cleanup, list):
        return any(_cleanup_has(value, receipt) for value in cleanup)
    return False


def load_cleanup() -> tuple[Any, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    verified = value.get("verified") is True or value.get("executables_verified_before_removal") is True
    return value, verified


def validate_binary(receipt: Any, label: str, cleanup: Any, cleanup_verified: bool) -> None:
    require(isinstance(receipt, dict), f"{label} receipt is missing")
    path = artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is not None:
        return
    require(cleanup_verified, f"{label} is missing without cleanup verification")
    require(_cleanup_has(cleanup, receipt), f"{label} lacks an exact cleanup witness")


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and revision and all(c in HEX for c in revision),
            f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                f"{label} contains an invalid file digest")
    return {"revision": revision, "files": dict(files)}


def current_source_files() -> dict[str, str]:
    """Hash the live source census without binding the result to HEAD."""

    try:
        raw = subprocess.check_output(
            [
                "git",
                "ls-files",
                "-z",
                "--",
                "crates",
                "Cargo.toml",
                "clippy.toml",
                ".cargo/config.toml",
                "rust-toolchain.toml",
            ],
            cwd=ROOT,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read live source census: {error}")
    names = [name for name in raw.decode().split("\0") if name]
    result: dict[str, str] = {}
    for name in names:
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"live source file is missing: {name}")
        result[name] = sha256(path)
    return result


def load_disposition(
    before_source: dict[str, Any], after_source: dict[str, Any]
) -> dict[str, Any]:
    """Verify the rejected candidate archive and restored live source."""

    path = PACKET / "disposition.json"
    require(path.is_file(), "disposition.json is missing")
    disposition = read_json(path)
    require(isinstance(disposition, dict), "disposition.json is malformed")
    require(disposition.get("status") == "rejected", "candidate disposition is not rejected")
    require(disposition.get("production_change_retained") is False,
            "rejected candidate remains marked as retained")
    require(disposition.get("source_file") == PHYS_PKG,
            "disposition source file changed")
    before_archive = artifact_path(disposition.get("before_source"), "before source archive")
    candidate_archive = artifact_path(disposition.get("candidate_source"),
                                      "candidate source archive")
    restored_receipt = disposition.get("restored_source")
    restored_path = artifact_path(restored_receipt, "restored source census")
    require(sha256(before_archive) == before_source["files"][PHYS_PKG],
            "before source archive does not match before census")
    require(sha256(candidate_archive) == after_source["files"][PHYS_PKG],
            "candidate source archive does not match after census")
    restored = source_manifest(read_json(restored_path), "restored source census")
    require(restored["files"] == before_source["files"],
            "restored source census differs from before census")
    live_files = current_source_files()
    require(live_files == before_source["files"],
            "live source files do not match before census after rejection")
    return {
        "status": "rejected",
        "production_change_retained": False,
        "source_file": PHYS_PKG,
        "before_source": {
            "path": rel(before_archive),
            "bytes": before_archive.stat().st_size,
            "sha256": sha256(before_archive),
        },
        "candidate_source": {
            "path": rel(candidate_archive),
            "bytes": candidate_archive.stat().st_size,
            "sha256": sha256(candidate_archive),
        },
        "restored_source": {
            "path": rel(restored_path),
            "bytes": restored_path.stat().st_size,
            "sha256": sha256(restored_path),
        },
        "live_source_files_match_before": True,
    }


def _probe_files() -> set[str]:
    root = PACKET / "probe-src"
    require(root.is_dir(), "probe source directory is missing")
    return {
        str(path.relative_to(PACKET))
        for path in root.rglob("*")
        if path.is_file() and path.name not in {"Cargo.lock", "Cargo.toml"}
    }


def load_builds(plan: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], Any, bool]:
    cleanup, cleanup_verified = load_cleanup()
    builds: dict[str, dict[str, Any]] = {}
    for leg in LEGS:
        directory = PACKET / f"build-{leg}"
        build = read_json(directory / "build.json")
        require(isinstance(build, dict), f"{leg} build manifest is malformed")
        source_receipt = build.get("source")
        source_path = artifact_path(source_receipt, f"{leg} build source")
        source = source_manifest(read_json(source_path), f"{leg} build source")
        probe = build.get("probe")
        require(isinstance(probe, dict), f"{leg} probe manifest is missing")
        require(set(probe) == _probe_files(), f"{leg} probe file inventory changed")
        for name, digest in probe.items():
            require(is_sha(digest), f"{leg} probe digest is invalid: {name}")
            path = PACKET / name
            require(path.is_file() and not path.is_symlink(), f"missing probe file {name}")
            require(sha256(path) == digest, f"{leg} probe file changed: {name}")
        lock = build.get("lock")
        lock_path = artifact_path(lock, f"{leg} probe lock")
        require(lock_path == (PACKET / "probe-src/Cargo.lock").resolve(),
                f"{leg} lock receipt is relocated outside probe packet")
        require(sha256(lock_path) == lock["sha256"], f"{leg} lock digest changed")
        require((PACKET / "probe-src/Cargo.toml").is_file(), f"{leg} generated probe manifest missing")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation"},
                f"{leg} binary map changed")
        for name, receipt in binaries.items():
            validate_binary(receipt, f"{leg} {name} binary", cleanup, cleanup_verified)
        rows = build.get("rows")
        require(isinstance(rows, list) and len(rows) == 2, f"{leg} build rows are incomplete")
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"{leg} build row {index} failed")
            artifact_path(row.get("log"), f"{leg} build row {index} log")
        builds[leg] = {
            "manifest": build,
            "source": source,
            "source_path": source_path,
            "binaries": binaries,
            "lock": lock,
            "probe": probe,
        }
    before = builds["before"]["source"]["files"]
    after = builds["after"]["source"]["files"]
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    require(changed == [PHYS_PKG], f"source census changed outside {PHYS_PKG}: {changed}")
    require(builds["before"]["lock"]["sha256"] == builds["after"]["lock"]["sha256"],
            "probe lock changed between before and after")
    require(builds["before"]["probe"] == builds["after"]["probe"],
            "probe source changed between before and after")
    for lane in ("native", "allocation", "qualification"):
        path = PACKET / lane / "source.json"
        if path.is_file():
            value = source_manifest(read_json(path), f"{lane} source census")
            expected = builds["before"]["source"] if lane == "qualification" else builds["after"]["source"]
            require(value == expected, f"{lane} source census differs from its build leg")
    return builds, cleanup, cleanup_verified


def check_quality(after_source: dict[str, Any]) -> dict[str, Any] | None:
    path = PACKET / "quality.json"
    if not path.is_file():
        return None
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is malformed")
    source_path = artifact_path(quality.get("source"), "quality source")
    require(source_manifest(read_json(source_path), "quality source") == after_source,
            "quality source census differs from after build")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate count changed")
    packages = ["-p", "litchi-opc", "-p", "litchi-xlsx"]
    commands = [
        ["cargo", "fmt", *packages, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", *packages, "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", *packages, "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", *packages, "--all-features", "--lib", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", *packages, "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    require([row.get("command") for row in rows] == commands,
            "quality gate commands changed")
    environment = quality.get("environment", {})
    for key, expected in {"CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
                          "CARGO_PROFILE_DEV_DEBUG": "0", "RUSTDOCFLAGS": "-D warnings"}.items():
        require(environment.get(key) == expected, f"quality environment changed: {key}")
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        log = artifact_path(row.get("log"), f"quality gate {index} log")
        logs.append(rel(log))
    return {"gates": len(rows), "logs": logs, "source": rel(source_path)}


def expected_jobs(plan: dict[str, Any], lane: str, *, mode: str | None = None) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    mode = mode or lane
    if lane == "qualification":
        schedule = [("before",)]
        blocks = 1
    else:
        blocks = plan[mode]["blocks"]
        schedule = [tuple(plan["native"]["orders"][block]) for block in range(blocks)]
    for block in range(blocks):
        legs = schedule[0] if lane == "qualification" else schedule[block]
        for corpus in plan["corpora"]:
            for phase in plan["phases"]:
                for leg in legs:
                    jobs.append({
                        "lane": lane, "block": block, "corpus": corpus["id"],
                        "phase": phase, "leg": leg,
                        "samples": 1 if lane == "qualification" else plan[mode]["samples"],
                        "warmup": 0 if lane == "qualification" else plan[mode]["warmup"],
                    })
    return jobs


def validate_command(command: Any, job: dict[str, Any], binary: dict[str, Any],
                     corpus: dict[str, Any], report_path: Path) -> None:
    require(isinstance(command, list) and all(isinstance(x, str) for x in command),
            "capture command is malformed")
    require("taskset" in command and str(plan_cpu) in command,
            "capture command CPU binding changed")
    for value in (str(corpus["sheet"]), job["phase"], str(job["samples"]),
                  str(job["warmup"])):
        require(value in command, f"capture command is missing {value}")
    require(any(
        (resolve_path(value, packet_bound=True) == report_path)
        for value in command
        if value.endswith(".json") or value.endswith(".JSON")
    ), "capture command output path changed")
    require(any(
        Path(value).name == Path(binary["path"]).name for value in command
    ), "capture command binary changed")


plan_cpu = 12


def load_lane(
    plan: dict[str, Any],
    lane: str,
    builds: dict[str, dict[str, Any]],
    after_source: dict[str, Any],
    cleanup: Any,
    cleanup_verified: bool,
) -> list[dict[str, Any]]:
    directory = PACKET / lane
    require(directory.is_dir(), f"missing {lane} capture directory")
    mode = lane.removeprefix("small-")
    complete = read_json(directory / "complete.json")
    require(isinstance(complete, dict), f"{lane} complete marker is malformed")
    jobs = expected_jobs(plan, lane, mode=mode)
    require(complete.get("children") == len(jobs),
            f"{lane} completion child count changed")
    complete_source = artifact_path(complete.get("source"), f"{lane} complete source")
    require(source_manifest(read_json(complete_source), f"{lane} complete source") == after_source,
            f"{lane} complete source differs from after source census")
    rows_path = artifact_path(complete.get("receipts"), f"{lane} receipts")
    rows = read_json(rows_path)
    require(isinstance(rows, list), f"{lane} receipts are not a list")
    require(len(rows) == len(jobs), f"{lane} receipt count changed")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"{lane} child {index}"
        require(isinstance(row, dict), f"{label} is not an object")
        for key in ("lane", "block", "corpus", "phase", "leg"):
            require(row.get(key) == job[key], f"{label} {key} identity changed")
        key = (job["block"], job["corpus"], job["phase"], job["leg"])
        require(key not in seen, f"duplicate {lane} identity: {key}")
        seen.add(key)
        require(row.get("exit_code") == 0, f"{label} failed")
        corpus = corpora(plan)[job["corpus"]]
        binary_kind = "native" if mode == "native" else "allocation"
        binary = row.get("binary")
        require(isinstance(binary, dict), f"{label} binary receipt is missing")
        expected_binary = builds[job["leg"]]["binaries"][binary_kind]
        require(binary.get("bytes") == expected_binary.get("bytes")
                and binary.get("sha256") == expected_binary.get("sha256"),
                f"{label} binary does not match {job['leg']} build")
        validate_binary(binary, f"{label} binary", cleanup, cleanup_verified)
        report_path = artifact_path(row.get("report"), f"{label} report")
        log_path = artifact_path(row.get("log"), f"{label} log")
        rss_path = artifact_path(row.get("rss"), f"{label} RSS")
        try:
            rss_text = rss_path.read_text().strip()
        except OSError as error:
            fail(f"{label} RSS is unreadable: {error}")
        require(rss_text.isdigit(), f"{label} RSS is not an integer")
        rss = int(rss_text)
        report = read_json(report_path)
        require(isinstance(report, dict), f"{label} report is not an object")
        validate_command(row.get("command"), job, expected_binary, corpus, report_path)
        outcome = validate_report(report, job, corpus, binary_kind, builds, label)
        entries.append({
            "identity": job, "row": row, "report": report, "report_path": report_path,
            "report_sha256": sha256(report_path), "log": rel(log_path), "rss_kib": rss,
            "stats": outcome["stats"], "allocation": outcome.get("allocation"),
            "published": outcome["published"],
        })
    require({(item["identity"]["block"], item["identity"]["corpus"],
               item["identity"]["phase"], item["identity"]["leg"]) for item in entries}
            == set(seen), f"{lane} schedule is incomplete")
    return entries


def nearest_rank(values: Iterable[int | float], percentile: int) -> int | float:
    vector = sorted(values)
    require(vector, "cannot compute percentile of an empty vector")
    rank = max(1, math.ceil(percentile * len(vector) / 100))
    return vector[rank - 1]


def stats(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        finite_number(value, f"metric[{index}]")
    return {
        "count": len(vector),
        "min": min(vector),
        "p50": nearest_rank(vector, 50),
        "mean": statistics.mean(vector),
        "p95": nearest_rank(vector, 95),
        "p99": nearest_rank(vector, 99),
        "max": max(vector),
    }


def spread(values: Iterable[int | float]) -> float:
    vector = [float(value) for value in values]
    require(vector and all(value >= 0 and math.isfinite(value) for value in vector),
            "spread vector is invalid")
    low, high = min(vector), max(vector)
    if low == 0:
        return 0.0 if high == 0 else float("inf")
    return (high - low) * 100.0 / low


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    result = stats(vector)
    result["values"] = vector
    result["spread_percent"] = spread(vector)
    result["flag_over_5_percent"] = result["spread_percent"] > 5.0
    return result


def timing_scope(phase: str) -> str:
    return {
        "open": "Workbook::open(source) only",
        "edit": "Workbook::edit, sheet(sheet).set(A1, marker), and commit only; open is outside the clock",
        "save": "Workbook::save_with_durability(destination, NoSync) only; open and edit are outside the clock",
        "lifecycle": "Workbook::open, semantic A1 edit and commit, then save_with_durability(destination, NoSync)",
    }[phase]


def previous_output(corpus: dict[str, Any]) -> tuple[int, str]:
    source = Path(corpus["source"])
    if not source.is_absolute():
        source = ROOT / source
    previous = source.parent / "no-sync.xlsx"
    require(previous.is_file() and not previous.is_symlink(),
            f"0778 no-sync output is missing for {corpus['id']}")
    return previous.stat().st_size, sha256(previous)


def validate_report(
    report: dict[str, Any],
    job: dict[str, Any],
    corpus: dict[str, Any],
    binary_kind: str,
    builds: dict[str, dict[str, Any]],
    label: str,
) -> dict[str, Any]:
    require(report.get("schema") == REPORT_SCHEMA, f"{label} report schema changed")
    require(report.get("tool") == REPORT_TOOL, f"{label} report tool changed")
    identity = report.get("source")
    require(isinstance(identity, dict), f"{label} source identity is missing")
    source_path = Path(corpus["source"])
    if not source_path.is_absolute():
        source_path = ROOT / source_path
    require(source_path.is_file(), f"{label} fixture is missing")
    require(identity.get("bytes") == corpus["bytes"] and identity.get("sha256") == corpus["sha256"],
            f"{label} report source identity changed")
    require(source_path.stat().st_size == corpus["bytes"] and sha256(source_path) == corpus["sha256"],
            f"{label} fixture identity changed")
    require(report.get("sheet") == corpus["sheet"], f"{label} sheet changed")
    require(report.get("address") == "A1" and report.get("marker") == MARKER,
            f"{label} marker identity changed")
    require(report.get("phase") == job["phase"], f"{label} phase changed")
    require(report.get("timing_scope") == timing_scope(job["phase"]),
            f"{label} timing scope changed")
    uses_output = job["phase"] in ("save", "lifecycle")
    require(report.get("durability") == ("no-sync" if uses_output else None),
            f"{label} durability changed")
    require(report.get("warmup") == job["warmup"]
            and report.get("samples_requested") == job["samples"],
            f"{label} sample configuration changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == job["samples"],
            f"{label} sample count changed")
    allocator = report.get("allocator")
    require(isinstance(allocator, dict), f"{label} allocator identity is missing")
    expected_allocator = (
        ("litchi-xlsx-allocation-probe-0779-alloc", "system_allocator_operation_scoped",
         "CountingSystemAllocator(std::alloc::System)", "serialized_region_peak_v3")
        if binary_kind == "allocation" else
        ("litchi-xlsx-allocation-probe-0779", "none", "Rust system allocator", None)
    )
    require(tuple(allocator.get(key) for key in
                  ("binary", "instrumentation", "allocator", "counter_revision"))
            == expected_allocator, f"{label} allocator identity changed")
    elapsed: list[int] = []
    allocations: dict[str, list[int]] = {field: [] for field in ALLOCATION_FIELDS}
    published: list[tuple[int, str]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label} sample {index} identity changed")
        elapsed_ns = sample.get("elapsed_ns")
        nonnegative_int(elapsed_ns, f"{label} elapsed_ns[{index}]")
        elapsed.append(elapsed_ns)
        verification = sample.get("verification")
        if verification is not None:
            require(isinstance(verification, dict) and verification.get("reopened") is True
                    and verification.get("sheet") == corpus["sheet"]
                    and verification.get("address") == "A1"
                    and verification.get("expected") == MARKER
                    and verification.get("actual") == MARKER
                    and verification.get("marker_matches") is True,
                    f"{label} marker verification changed at sample {index}")
        if uses_output:
            nonnegative_int(sample.get("published_bytes"), f"{label} published bytes[{index}]")
            digest = sample.get("published_sha256")
            require(is_sha(digest), f"{label} published SHA missing at sample {index}")
            published.append((sample["published_bytes"], digest))
            require(verification is not None, f"{label} output marker verification missing")
        else:
            require(sample.get("published_bytes") is None
                    and sample.get("published_sha256") is None,
                    f"{label} unexpected published output")
        allocation = sample.get("allocation")
        if binary_kind == "native":
            require(allocation is None, f"{label} native report contains allocation metrics")
        else:
            require(isinstance(allocation, dict) and allocation.get("status") == "measured"
                    and allocation.get("scope") == "operation_global_system_allocator",
                    f"{label} allocation sample is not measured")
            for field in ALLOCATION_FIELDS:
                value = allocation.get(field)
                nonnegative_int(value, f"{label} allocation {field}[{index}]")
                allocations[field].append(value)
            before = allocation["live_bytes_before"]
            after = allocation["live_bytes_after"]
            require(after == before + allocation["allocated_bytes"] - allocation["deallocated_bytes"],
                    f"{label} live-byte accounting changed at sample {index}")
            require(allocation["region_peak_live_bytes"] >= before
                    and allocation["region_peak_live_bytes"] >= after,
                    f"{label} region peak is below entry/exit at sample {index}")
            require(allocation["peak_live_bytes_before"] >= before
                    and allocation["peak_live_bytes_after"] >= after
                    and allocation["peak_live_bytes_after"] >= allocation["peak_live_bytes_before"]
                    and allocation["peak_live_bytes_after"] >= allocation["region_peak_live_bytes"],
                    f"{label} allocator peak ordering changed at sample {index}")
            require(allocation["failed_allocation_calls"] == 0,
                    f"{label} failed allocation observed")
    if uses_output:
        require(len(set(published)) == 1, f"{label} output digest varied within report")
        expected_size, expected_digest = previous_output(corpus)
        require(published[0] == (expected_size, expected_digest),
                f"{label} output differs from 0778 no-sync artifact")
    else:
        published = []
    alloc_vectors = None
    if binary_kind == "allocation":
        allocations["net_live"] = [
            after - before for before, after in zip(
                allocations["live_bytes_before"], allocations["live_bytes_after"])
        ]
        allocations["peak_above_entry"] = [
            peak - before for peak, before in zip(
                allocations["region_peak_live_bytes"], allocations["live_bytes_before"])
        ]
        alloc_vectors = allocations
    return {
        "stats": stats(elapsed),
        "allocation": alloc_vectors,
        "published": [list(pair) for pair in published],
    }


def pair_ratio(before: float, after: float) -> dict[str, Any]:
    if before == 0:
        return {
            "before": before, "after": after, "ratio": None,
            "change_percent": None,
            "over_5_percent": after > 0,
        }
    ratio = after / before
    return {
        "before": before, "after": after, "ratio": ratio,
        "change_percent": (ratio - 1.0) * 100.0,
        "over_5_percent": ratio > 1.05,
    }


def paired(entries: list[dict[str, Any]], metrics: Iterable[str]) -> dict[str, Any]:
    by_key = {
        (item["identity"]["corpus"], item["identity"]["phase"],
         item["identity"]["leg"], item["identity"]["block"]): item
        for item in entries
    }
    output: dict[str, Any] = {}
    for corpus in sorted({item["identity"]["corpus"] for item in entries}):
        for phase in PHASES:
            blocks = sorted({item["identity"]["block"] for item in entries
                             if item["identity"]["corpus"] == corpus})
            if not blocks or any(
                (corpus, phase, leg, block) not in by_key
                for leg in ("before", "after") for block in blocks
            ):
                continue
            before = [by_key[(corpus, phase, "before", block)] for block in blocks]
            after = [by_key[(corpus, phase, "after", block)] for block in blocks]
            require(len(before) == len(after), f"paired rows incomplete: {corpus}/{phase}")
            values: dict[str, Any] = {}
            for metric in metrics:
                ratios = []
                per_block = []
                for left, right in zip(before, after):
                    if metric == "rss_kib":
                        left_value = left["rss_kib"]
                        right_value = right["rss_kib"]
                    elif left.get("allocation") is not None and metric in left["allocation"]:
                        left_value = stats(left["allocation"][metric])["p50"]
                        right_value = stats(right["allocation"][metric])["p50"]
                    else:
                        left_value = left["stats"][metric]
                        right_value = right["stats"][metric]
                    result = pair_ratio(float(left_value), float(right_value))
                    result["block"] = left["identity"]["block"]
                    per_block.append(result)
                    if result["ratio"] is not None:
                        ratios.append(result["ratio"])
                median = statistics.median(ratios) if ratios else None
                values[metric] = {
                    "by_block": per_block,
                    "ratio_median": median,
                    "change_percent_median": None if median is None else (median - 1.0) * 100.0,
                    "regression_over_5_percent": (
                        median is not None and median > 1.05
                    ) or any(item["over_5_percent"] for item in per_block),
                }
            output[f"{corpus}/{phase}"] = {
                "corpus": corpus,
                "phase": phase,
                "blocks": len(before),
                "metrics": values,
                "comparison": "after/before paired by capture block; descriptive regression flag",
            }
    return output


def native_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["corpus"], identity["phase"], identity["leg"]), []).append(item)
    for key in sorted(grouped):
        values = grouped[key]
        elapsed_distributions = {}
        for metric in METRICS:
            value = distribution(item["stats"][metric] for item in values)
            elapsed_distributions[metric] = value
            if value["flag_over_5_percent"]:
                spread_flags.append({"group": list(key), "metric": metric,
                                     "spread_percent": value["spread_percent"]})
        rss_distribution = distribution(item["rss_kib"] for item in values)
        if rss_distribution["flag_over_5_percent"]:
            spread_flags.append({"group": list(key), "metric": "rss_kib",
                                 "spread_percent": rss_distribution["spread_percent"]})
        groups["/".join(key)] = {
            "corpus": key[0], "phase": key[1], "leg": key[2],
            "processes": [
                {"block": item["identity"]["block"], "stats": item["stats"],
                 "rss_kib": item["rss_kib"], "report": rel(item["report_path"]),
                 "report_sha256": item["report_sha256"]}
                for item in sorted(values, key=lambda item: item["identity"]["block"])
            ],
            "elapsed_distribution_across_processes": elapsed_distributions,
            "per_process_rss_distribution": rss_distribution,
        }
    paired_values = paired(entries, (*METRICS, "rss_kib"))
    regression_flags = []
    for key, group in paired_values.items():
        for metric, value in group["metrics"].items():
            if value["regression_over_5_percent"]:
                regression_flags.append({
                    "group": key, "metric": metric,
                    "ratio_median": value["ratio_median"],
                    "change_percent_median": value["change_percent_median"],
                })
    return {
        "groups": groups,
        "spread_flags_over_5_percent": spread_flags,
        "regression_flags_over_5_percent": regression_flags,
        "paired_by_block_before_after": paired_values,
        "rss_is_separate_from_elapsed": True,
        "allocation_metrics_present": False,
    }


def allocation_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    fields = (*ALLOCATION_FIELDS, "net_live", "peak_above_entry")
    groups: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["corpus"], identity["phase"], identity["leg"]), []).append(item)
    for key in sorted(grouped):
        values = grouped[key]
        metrics: dict[str, Any] = {}
        for field in fields:
            per_block = [
                {
                    "block": item["identity"]["block"],
                    "values": item["allocation"][field],
                    "stats": stats(item["allocation"][field]),
                }
                for item in sorted(values, key=lambda item: item["identity"]["block"])
            ]
            repeats = [item["stats"]["p50"] for item in per_block]
            value = {
                "per_block": per_block,
                "repeat_p50_values": repeats,
                "spread_percent": spread(repeats),
                "flag_over_5_percent": spread(repeats) > 5.0,
            }
            metrics[field] = value
            if value["flag_over_5_percent"]:
                spread_flags.append({"group": list(key), "metric": field,
                                     "spread_percent": value["spread_percent"]})
        groups["/".join(key)] = {
            "corpus": key[0], "phase": key[1], "leg": key[2],
            "blocks": len(values), "metrics": metrics,
            "elapsed_not_mixed": True,
            "region_peak_and_entry_are_separate": True,
        }
    paired_values = paired(entries, fields)
    regression_flags = []
    for key, group in paired_values.items():
        for metric, value in group["metrics"].items():
            if value["regression_over_5_percent"]:
                regression_flags.append({
                    "group": key, "metric": metric,
                    "ratio_median": value["ratio_median"],
                    "change_percent_median": value["change_percent_median"],
                })
    return {
        "groups": groups,
        "fields": list(fields),
        "spread_flags_over_5_percent": spread_flags,
        "regression_flags_over_5_percent": regression_flags,
        "paired_by_block_before_after": paired_values,
        "allocation_is_separate_from_elapsed": True,
    }


def check_qualification(entries: list[dict[str, Any]]) -> dict[str, Any]:
    require(all(item["identity"]["leg"] == "before" for item in entries),
            "qualification contains an after leg")
    return {
        "children": len(entries),
        "rows": [
            {
                "corpus": item["identity"]["corpus"],
                "phase": item["identity"]["phase"],
                "report": rel(item["report_path"]),
                "report_sha256": item["report_sha256"],
                "published": item["published"],
            }
            for item in entries
        ],
        "before_source_only": True,
    }


def load_heaptrack() -> dict[str, Any]:
    """Replay the independent Heaptrack role verifier and bind its report."""

    analysis_path = PACKET / "heaptrack-analysis.json"
    runner_path = PACKET / "heaptrack_analysis.py"
    require(analysis_path.is_file(), "heaptrack-analysis.json is missing")
    require(runner_path.is_file() and not runner_path.is_symlink(),
            "heaptrack analysis runner is missing")
    stored = read_json(analysis_path)
    require(isinstance(stored, dict)
            and stored.get("schema") == "litchi.xlsx.heaptrack-analysis.v1"
            and stored.get("status") == "pass",
            "heaptrack analysis status or schema changed")
    roles = stored.get("roles")
    require(isinstance(roles, dict) and set(roles) == {"before", "after"},
            "heaptrack role set changed")
    spec = importlib.util.spec_from_file_location("_0779_heaptrack_analysis", runner_path)
    require(spec is not None and spec.loader is not None,
            "heaptrack analysis runner cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    fresh = module.analyze(PACKET.resolve())
    require(stored == fresh, "heaptrack-analysis.json does not replay")
    return {
        "analysis": stored,
        "analysis_receipt": {
            "path": rel(analysis_path),
            "bytes": analysis_path.stat().st_size,
            "sha256": sha256(analysis_path),
        },
        "runner_receipt": {
            "path": rel(runner_path),
            "bytes": runner_path.stat().st_size,
            "sha256": sha256(runner_path),
        },
    }


def check_optional_seal() -> bool:
    for name in ("seal.json", "final-seal.json"):
        path = PACKET / name
        if not path.is_file():
            continue
        value = read_json(path)
        require(isinstance(value, dict) and isinstance(value.get("files"), dict),
                f"{name} is malformed")
        actual = {
            rel(item): sha256(item)
            for item in PACKET.rglob("*")
            if item.is_file() and item.name not in {"seal.json", "final-seal.json"}
        }
        require(value["files"] == actual, f"{name} file inventory is stale")
        return True
    return False


def analyze() -> dict[str, Any]:
    plan = load_plan()
    global plan_cpu
    plan_cpu = plan.get("cpu", 12)
    require(plan_cpu == 12, "capture CPU changed")
    builds, cleanup, cleanup_verified = load_builds(plan)
    after_source = builds["after"]["source"]
    disposition = load_disposition(builds["before"]["source"], after_source)
    quality = check_quality(after_source)
    native_entries = load_lane(plan, "native", builds, after_source, cleanup, cleanup_verified)
    allocation_entries = load_lane(
        plan, "allocation", builds, after_source, cleanup, cleanup_verified
    )
    qualification_entries = load_lane(
        plan, "qualification", builds, builds["before"]["source"], cleanup, cleanup_verified
    )
    require(len(native_entries) == 96, "native cardinality changed")
    require(len(allocation_entries) == 32, "allocation cardinality changed")
    require(len(qualification_entries) == 8, "qualification cardinality changed")
    small_plan = load_small_plan()
    small_native_entries = load_lane(
        small_plan, "small-native", builds, after_source, cleanup, cleanup_verified
    )
    small_allocation_entries = load_lane(
        small_plan, "small-allocation", builds, after_source, cleanup, cleanup_verified
    )
    require(len(small_native_entries) == 24, "small native cardinality changed")
    require(len(small_allocation_entries) == 8, "small allocation cardinality changed")
    native = native_analysis(native_entries)
    allocation = allocation_analysis(allocation_entries)
    small_native = native_analysis(small_native_entries)
    small_allocation = allocation_analysis(small_allocation_entries)
    qualification = check_qualification(qualification_entries)
    heaptrack = load_heaptrack()
    check_optional_seal()
    return {
        "schema": "litchi-0779-xlsx-allocation-analysis-v1",
        "plan_schema": plan["schema"],
        "base": plan["base"],
        "quality": quality,
        "source": {
            "before": builds["before"]["source"],
            "after": builds["after"]["source"],
            "changed_files": [PHYS_PKG],
        },
        "disposition": disposition,
        "native": {
            "children": len(native_entries),
            "blocks": plan["native"]["blocks"],
            "samples": plan["native"]["samples"],
            "warmup": plan["native"]["warmup"],
            "analysis": native,
            "receipts": [
                {
                    "corpus": item["identity"]["corpus"],
                    "phase": item["identity"]["phase"],
                    "block": item["identity"]["block"],
                    "leg": item["identity"]["leg"],
                    "report": rel(item["report_path"]),
                    "report_sha256": item["report_sha256"],
                }
                for item in native_entries
            ],
        },
        "allocation": {
            "children": len(allocation_entries),
            "blocks": plan["allocation"]["blocks"],
            "samples": plan["allocation"]["samples"],
            "warmup": plan["allocation"]["warmup"],
            "analysis": allocation,
            "receipts": [
                {
                    "corpus": item["identity"]["corpus"],
                    "phase": item["identity"]["phase"],
                    "block": item["identity"]["block"],
                    "leg": item["identity"]["leg"],
                    "report": rel(item["report_path"]),
                    "report_sha256": item["report_sha256"],
                }
                for item in allocation_entries
            ],
        },
        "qualification": qualification,
        "heaptrack": heaptrack,
        "small_controls": {
            "plan_schema": small_plan["schema"],
            "native": {
                "children": len(small_native_entries),
                "blocks": small_plan["native"]["blocks"],
                "samples": small_plan["native"]["samples"],
                "warmup": small_plan["native"]["warmup"],
                "analysis": small_native,
                "receipts": [
                    {
                        "corpus": item["identity"]["corpus"],
                        "phase": item["identity"]["phase"],
                        "block": item["identity"]["block"],
                        "leg": item["identity"]["leg"],
                        "report": rel(item["report_path"]),
                        "report_sha256": item["report_sha256"],
                    }
                    for item in small_native_entries
                ],
            },
            "allocation": {
                "children": len(small_allocation_entries),
                "blocks": small_plan["allocation"]["blocks"],
                "samples": small_plan["allocation"]["samples"],
                "warmup": small_plan["allocation"]["warmup"],
                "analysis": small_allocation,
                "receipts": [
                    {
                        "corpus": item["identity"]["corpus"],
                        "phase": item["identity"]["phase"],
                        "block": item["identity"]["block"],
                        "leg": item["identity"]["leg"],
                        "report": rel(item["report_path"]),
                        "report_sha256": item["report_sha256"],
                    }
                    for item in small_allocation_entries
                ],
            },
        },
        "verification": {
            "source_binary_probe_lock_fixture_receipts_checked": True,
            "source_change_allowlist": [PHYS_PKG],
            "native_has_no_allocation_metrics": True,
            "published_output_digest_before_after_checked": True,
            "published_output_digest_matches_0778_no_sync": True,
            "marker_verification_checked": True,
            "small_open_controls_separate_from_primary_matrix": True,
            "cleanup_binary_witness_required_when_missing": True,
            "cleanup_witness_checked": True,
            "optional_file_inventory_seal_checked": True,
        },
        "limits": [
            "Timing and allocation comparisons are descriptive before/after evidence.",
            "RSS is a whole-process /usr/bin/time gauge including setup and verification.",
            "Allocation region fields are not latency, physical-copy, or phase-peak estimates.",
            "No cold-cache, device-floor, concurrency, or causal performance claim is made.",
        ],
    }


if __name__ == "__main__":
    result = analyze()
    output = PACKET / "analysis.json"
    require(not output.exists(), "refusing to overwrite analysis.json")
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({
        "native": result["native"]["children"],
        "allocation": result["allocation"]["children"],
        "qualification": result["qualification"]["children"],
    }, sort_keys=True))
