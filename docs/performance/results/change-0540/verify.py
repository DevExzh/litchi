"""Read-only verifier for the 0540 XLSX planning attribution bundle."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
import re
import subprocess
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
FROZEN = HERE / "frozen-inputs.json"
BASE = HERE / "baseline"
PRIOR = HERE.parent / "change-0539"
PRIOR_MANIFEST = PRIOR / "baseline" / "source-manifest.json"
PRIOR_SEAL = PRIOR / "SHA256SUMS"
ANALYZER = HERE / "analyze_planning.py"
BOUNDARY_ANALYZER = HERE / "analyze_boundary.py"
TARGET = Path("/home/zhuhe/litchi-goal-0540-target")
OWNED = [str(TARGET)]
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SHAPES = ["medium", "dense-sparse"]
OWNERS = [
    "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets",
    "litchi_xlsx::cell_values::snapshot::MultiSnapshot::load_source_backed",
]
PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
LIFECYCLE = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALLOCATIONS = (
    "plan_allocation_metrics",
    "commit_allocation_metrics",
    "publication_allocation_metrics",
)
QUALITY_NAMES = (
    "final-quality-boundaries",
    "final-quality-check",
    "final-quality-clippy",
    "final-quality-fmt",
    "final-quality-rustdoc",
    "final-quality-tests",
)


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


class Pending(EvidenceError):
    """Evidence is not complete yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def need(path: Path, label: str | None = None) -> Path:
    name = label or str(path)
    if not path.exists():
        raise Pending(f"{name} is missing")
    require(not path.is_symlink(), f"{name} is a symlink")
    return path


def read(path: Path) -> Any:
    need(path)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def valid_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def timestamp(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is invalid") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    start = timestamp(value.get("start_utc"), f"{label}.start_utc")
    end = timestamp(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds > 0 and end > start,
            f"{label} interval is invalid")
    return start, end


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside bundle: {path}") from error


def current_sources() -> set[str]:
    tracked = subprocess.check_output([
        "git", "ls-files", "-z", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ], cwd=REPO).split(b"\0")
    untracked = subprocess.check_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ], cwd=REPO).split(b"\0")
    names = {item.decode() for item in tracked if item}
    names.update(item.decode() for item in untracked if item and item.endswith(b".rs"))
    return {name for name in names if (REPO / name).is_file()}


def source_manifest(path: Path) -> dict[str, str]:
    value = read(path)
    require(isinstance(value, dict) and value, f"{path} is not a source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe(name, f"{path} source path")
        require((name.startswith("crates/") or name.startswith("tools/perf-baseline/")
                 or name.startswith(".cargo/") or name in
                 ("Cargo.toml", "Cargo.lock", "rust-toolchain.toml"))
                and valid_digest(digest), f"{path} has an invalid source entry: {name}")
        require(name not in result, f"{path} repeats {name}")
        result[name] = digest
    return result


def frozen_inputs() -> tuple[dict[str, Any], dt.datetime]:
    frozen = read(FROZEN)
    require(isinstance(frozen, dict) and isinstance(frozen.get("files"), dict),
            "frozen input envelope differs")
    created = timestamp(frozen.get("created_utc"), "frozen-inputs.created_utc")
    files = frozen["files"]
    require(set(files) == {"plan.json", "run.py", "analyze_planning.py", "adr-manifest.json"},
            "frozen input inventory differs")
    for name, digest in files.items():
        require(valid_digest(digest) and sha(HERE / name) == digest,
                f"frozen input hash differs: {name}")
    return frozen, created


def plan() -> dict[str, Any]:
    _, _ = frozen_inputs()
    value = read(PLAN)
    require(value.get("status") == "frozen-before-build-and-capture"
            and isinstance(value.get("revision"), str)
            and re.fullmatch(r"[0-9a-f]{40}", value["revision"]),
            "plan is not frozen")
    require(value.get("priority") == "OLE2/OOXML active; ODF deferred; iWork excluded",
            "plan priority differs")
    primary, profile = value.get("primary"), value.get("profile")
    require(isinstance(primary, dict) and primary.get("case") == CASE
            and primary.get("shapes") == SHAPES, "primary plan differs")
    require(isinstance(profile, dict) and profile.get("shapes") == SHAPES
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1
            and profile.get("owner_candidates") == OWNERS,
            "profile plan differs")
    require(value.get("candidate_source_roots") == ["crates/litchi-xlsx/"]
            and value.get("capture_lanes") == ["build-normal", "profile"]
            and value.get("owned_paths") == OWNED,
            "scope or owned paths differ")
    require(not any(key in value for key in ("candidate", "guards", "allocation", "hardware")),
            "0540 unexpectedly contains a numerical or candidate lane")
    require("attribution" in str(value.get("admission", "")).lower()
            and "without changing runtime" in str(value.get("hypothesis", "")).lower(),
            "attribution-only scope is not frozen")
    return value


def validate_environment() -> dict[str, Any]:
    value = read(HERE / "environment.json")
    require(isinstance(value, dict) and value.get("scope") ==
            "Process-visible environment only; no host-wide quiescence claim.",
            "environment scope differs")
    require(isinstance(value.get("created_utc"), str), "environment timestamp is missing")
    timestamp(value["created_utc"], "environment.created_utc")
    require(isinstance(value.get("rustc"), str) and "rustc 1.95.0" in value["rustc"],
            "environment rustc evidence differs")
    require(isinstance(value.get("valgrind"), str) and value["valgrind"].startswith("valgrind-"),
            "environment valgrind evidence differs")
    return {"sha256": sha(HERE / "environment.json"), "scope": value["scope"]}


def validate_seal(path: Path, root: Path) -> int:
    expected: dict[str, str] = {}
    for line in need(path).read_text(encoding="utf-8").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), f"malformed seal line: {line}")
        name = safe(fields[1], "seal path")
        require(name != path.name and name not in expected, "seal repeats or contains itself")
        target = root / name
        require(target.is_file() and not target.is_symlink(), f"sealed file is missing: {name}")
        expected[name] = fields[0]
    actual = {
        item.relative_to(root).as_posix(): sha(item)
        for item in root.rglob("*")
        if item.is_file() and not item.is_symlink() and item != path
    }
    require(actual == expected and not any(item.is_symlink() for item in root.rglob("*")),
            "sealed inventory is not exact")
    return len(expected)


def validate_source() -> dict[str, Any]:
    value = plan()
    current = source_manifest(BASE / "source-manifest.json")
    prior = source_manifest(PRIOR_MANIFEST)
    require((BASE / "source-manifest.json").read_bytes() == PRIOR_MANIFEST.read_bytes()
            and current == prior, "baseline is not byte-equal to sealed 0539 baseline")
    require((BASE / "source.patch").read_bytes() == b"", "baseline source patch is not empty")
    require(current_sources() == set(current), "current source inventory differs")
    for name, digest in current.items():
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"source hash differs: {name}")
    adr = read(HERE / "adr-manifest.json").get("files")
    require(isinstance(adr, dict) and adr, "ADR manifest is missing")
    for name, digest in adr.items():
        safe(name, "ADR path")
        require(valid_digest(digest) and sha(REPO / name) == digest,
                f"ADR hash differs: {name}")
    seal_lines = {
        line.split("  ", 1)[1]: line.split("  ", 1)[0]
        for line in PRIOR_SEAL.read_text(encoding="utf-8").splitlines()
        if "  " in line
    }
    require(seal_lines.get("baseline/source-manifest.json") == sha(PRIOR_MANIFEST),
            "0539 seal does not bind its baseline source manifest")
    return {"manifest_sha256": sha(BASE / "source-manifest.json"),
            "manifest_entries": len(current), "prior_sealed_baseline_equal": True,
            "priority": value["priority"]}


def common(path: Path, source_sha: str, binary_sha: str | None = None) -> dict[str, Any]:
    value = read(path)
    start, end = interval(value, relative(path))
    require(value.get("exit_code") == 0
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("source_manifest_sha256") == source_sha
            and value.get("working_source_manifest_sha256") == source_sha,
            f"{path} receipt binding differs")
    require(value.get("environment", {}).get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{path} TMPDIR differs")
    if binary_sha is not None:
        require(value.get("binary_sha256") == binary_sha, f"{path} binary binding differs")
    return {"path": relative(path), "start": start, "end": end,
            "sha256": sha(path), "value": value}


def artifacts(value: dict[str, Any], folder: Path, label: str,
              expected: set[str] | None = None) -> None:
    entries = value.get("artifacts")
    require(isinstance(entries, dict), f"{label}: artifacts is not an object")
    if expected is not None:
        require(set(entries) == expected,
                f"{label}: artifact inventory differs: {sorted(set(entries) ^ expected)}")
    for filename, digest in entries.items():
        safe(filename, f"{label} artifact")
        target = folder / filename
        require(Path(filename).name == filename and target.is_file()
                and not target.is_symlink() and valid_digest(digest)
                and sha(target) == digest, f"{label}: artifact hash differs: {filename}")


def binary_identity() -> dict[str, Any]:
    source_sha = sha(BASE / "source-manifest.json")
    build_path = BASE / "build-normal.receipt.json"
    row = common(build_path, source_sha)
    _, frozen_created = frozen_inputs()
    require(frozen_created < row["start"],
            "normal build started before frozen inputs were recorded")
    command = [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline", "--target-dir", str(TARGET),
    ]
    require(row["value"].get("binary_sha256") is None
            and row["value"].get("command") == command,
            "build-normal command differs")
    artifacts(row["value"], BASE, "build-normal",
              {"build-normal.stdout", "build-normal.stderr"})
    identity = read(BASE / "binary-normal.json")
    expected_path = TARGET / "retained-binaries" / "baseline-normal"
    require(identity.get("path") == str(expected_path)
            and valid_digest(identity.get("sha256"))
            and isinstance(identity.get("bytes"), int) and identity["bytes"] > 0
            and identity.get("source_manifest_sha256") == source_sha
            and identity.get("build_receipt_sha256") == row["sha256"],
            "binary identity differs")
    if expected_path.exists():
        require(expected_path.is_file() and not expected_path.is_symlink()
                and sha(expected_path) == identity["sha256"]
                and expected_path.stat().st_size == identity["bytes"],
                "binary custody differs")
    else:
        cleanup = read(HERE / "cleanup.json")
        require(cleanup.get("owned_paths_absent") is True
                and cleanup.get("removed") == OWNED
                and cleanup.get("accessible_process_references") == [],
                "missing binary is not cleanup-bound")
        retained = cleanup.get("retained_binary_sha256_before_removal", {})
        require(retained.get(str(expected_path)) == identity["sha256"],
                "cleanup does not retain the binary hash")
    return {"receipt": row, "identity": identity,
            "binary_sha256": identity["sha256"], "binary_path": str(expected_path)}


def symbol(binary: dict[str, Any]) -> dict[str, Any]:
    source_sha = sha(BASE / "source-manifest.json")
    path = BASE / "symbols.receipt.json"
    row = common(path, source_sha, binary["binary_sha256"])
    value = row["value"]
    require(value.get("command") == ["nm", "-C", binary["binary_path"]],
            "symbols command differs")
    artifacts(value, BASE, "symbols", {"symbols.stdout", "symbols.stderr"})
    observation = read(HERE / "symbol-observation.json")
    candidates = observation.get("candidates")
    require(observation.get("plan_sha256") == sha(PLAN)
            and observation.get("binary_sha256") == binary["binary_sha256"]
            and observation.get("command") == ["nm", "-C", binary["binary_path"]]
            and isinstance(candidates, dict) and list(candidates) == OWNERS,
            "symbol observation binding differs")
    output = BASE / "symbols.stdout"
    require(observation.get("stdout_sha256") == sha(output)
            and observation.get("receipt_sha256") == row["sha256"],
            "symbol output custody differs")
    raw = output.read_text(encoding="utf-8")
    for owner, lines in candidates.items():
        require(isinstance(lines, list) and all(isinstance(line, str) for line in lines),
                f"symbol observations for {owner} are malformed")
        require(all(line in raw.splitlines() for line in lines),
                f"symbol observations for {owner} are absent from nm output")
    owner = observation.get("owner")
    require(owner in OWNERS and candidates[owner], "symbol owner is not present")
    first = next((name for name in OWNERS if candidates[name]), None)
    require(first == owner, "selected owner is not first exact nm match")
    return {"owner": owner, "first_match": first, "nm_sha256": sha(output),
            "receipt": row["sha256"]}


def helper_bindings() -> dict[str, Any]:
    """Bind the immutable helpers used by the planning analyzer itself."""
    value = read(HERE / "helper-binding.json")
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files, "helper binding inventory is missing")
    for name, digest in files.items():
        safe(name, "helper path")
        path = REPO / name
        require(valid_digest(digest) and path.is_file() and not path.is_symlink()
                and sha(path) == digest, f"helper hash differs: {name}")
    return {"files": len(files), "sha256": sha(HERE / "helper-binding.json")}


def profile_command(name: str, shape: str, owner: str, binary_path: str,
                    cpu: Any) -> list[str]:
    return [
        "taskset", "-c", str(cpu), "valgrind", "--tool=callgrind", "--collect-atstart=no",
        "--toggle-collect=" + owner, "--zero-before=" + owner,
        "--dump-after=" + owner,
        "--callgrind-out-file=" + str(BASE / (name + ".callgrind")), binary_path,
        "--warmup", "0", "--samples", "1", "--case", CASE,
        "--xlsx-cell-crud-shape", shape, "--json", str(BASE / (name + ".json")),
    ]


def phase_summary(raw: dict[str, Any], label: str) -> dict[str, Any]:
    result = raw["results"][0]
    elapsed = result["elapsed_ns"]
    elapsed_values = elapsed["samples"]
    xlsx = result["source"]["xlsx_cell_values"]
    values = {}
    for phase in PHASES:
        value = xlsx.get(phase)
        require(isinstance(value, list) and len(value) == 1
                and isinstance(value[0], int) and not isinstance(value[0], bool)
                and value[0] >= 0, f"{label}.{phase} is invalid")
        values[phase] = value[0]
    require(isinstance(elapsed_values, list) and len(elapsed_values) == 1
            and isinstance(elapsed_values[0], int) and elapsed_values[0] >= 0,
            f"{label}.elapsed_ns.samples is invalid")
    total = sum(values.values())
    require(total == elapsed_values[0], f"{label} phase sum differs from elapsed")
    return {**values, "sum_ns": total, "elapsed_ns": elapsed_values[0],
            "phase_sum_matches_elapsed": True}


def normal_allocation(raw: dict[str, Any], label: str) -> dict[str, Any]:
    xlsx = raw["results"][0]["source"]["xlsx_cell_values"]
    result = {}
    expected = {"status": "unavailable", "scope": "operation_global_system_allocator"}
    for name in ALLOCATIONS:
        values = xlsx.get(name)
        require(isinstance(values, list) and len(values) == 1 and values[0] == expected,
                f"{label}.{name} is not normal allocation unavailability")
        result[name] = expected
    return result


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def canonical(value: Any, label: str) -> Any:
    """Collapse one-sample vectors while retaining all static report fields."""
    if isinstance(value, list):
        require(len(value) == 1, f"{label} is not a one-sample vector")
        return canonical(value[0], f"{label}[0]")
    if isinstance(value, dict):
        return {key: canonical(item, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


def profile_static_identity(result: dict[str, Any], label: str) -> dict[str, Any]:
    source = result.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    retained = json.loads(json.dumps(source))
    xlsx = retained.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), f"{label}.source.xlsx_cell_values is not an object")
    for field in (*PHASES, "reopen_ns", *ALLOCATIONS):
        xlsx.pop(field, None)
    return {
        # Corpus metadata contains ordinary fixed-size lists (for example the
        # four worksheet member names), so preserve that object verbatim.
        "corpus": json.loads(json.dumps(result.get("corpus"))),
        "sink": canonical(result.get("sink"), f"{label}.sink"),
        "source": canonical(retained, f"{label}.source"),
        "output_sha256": result.get("output_sha256"),
    }


def validate_elapsed(result: dict[str, Any], label: str) -> int:
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}.elapsed_ns is not an object")
    samples = elapsed.get("samples")
    require(samples == sorted(samples) if isinstance(samples, list) else False,
            f"{label}.elapsed_ns.samples is malformed")
    require(isinstance(samples, list) and len(samples) == 1,
            f"{label}.elapsed_ns.samples is not one value")
    nonnegative_integer(samples[0], f"{label}.elapsed_ns.samples[0]")
    require(elapsed.get("unit") == "ns" and elapsed.get("sample_order") == [0],
            f"{label}.elapsed_ns metadata differs")
    sample = samples[0]
    require(all(elapsed.get(key) == sample for key in ("min", "p50", "p95", "p99", "max"))
            and elapsed.get("mean") == float(sample)
            and elapsed.get("standard_deviation") == 0.0,
            f"{label}.elapsed_ns statistics differ")
    confidence = elapsed.get("confidence_interval_95")
    require(isinstance(confidence, dict)
            and confidence.get("method") == "two-sided Student's t interval for the mean"
            and confidence.get("lower") == float(sample)
            and confidence.get("upper") == float(sample),
            f"{label}.elapsed_ns confidence interval differs")
    return sample


def profile_report(path: Path, name: str, shape: str, repeat: int,
                   binary: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    raw = read(path)
    label = name
    require(raw.get("schema_version") == 1, f"{label} schema version differs")
    tool = raw.get("tool")
    require(isinstance(tool, dict) and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release" and tool.get("instrumentation") == "none",
            f"{label} tool identity differs")
    identity = raw.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("path") == binary["binary_path"]
            and identity.get("binary_sha256") == binary["binary_sha256"]
            and identity.get("profile") == "release",
            f"{name} binary identity differs")
    environment = raw.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == p["revision"]
            and environment.get("cpu_affinity") == str(p["cpu"]),
            f"{label} environment identity differs")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [CASE]
            and configuration.get("xlsx_cell_crud_shapes") == [shape]
            and configuration.get("samples_per_case") == 1
            and configuration.get("warmup_iterations_per_case") == 0,
            f"{label} configuration differs")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1
            and isinstance(results[0], dict), f"{label} result matrix differs")
    result = results[0]
    require(result.get("case") == CASE, f"{label} result case differs")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == shape,
            f"{label} corpus identity differs")
    elapsed = validate_elapsed(result, label)
    sink = result.get("sink")
    require(isinstance(sink, dict), f"{label}.sink is not an object")
    for field in ("accepted_bytes", "write_calls", "largest_write"):
        nonnegative_integer(sink.get(field), f"{label}.sink.{field}")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}.sink buckets are missing")
    for field, value in buckets.items():
        nonnegative_integer(value, f"{label}.sink.write_size_buckets.{field}")
    require(sink["largest_write"] <= 65_536
            and sum(buckets.values()) == sink["write_calls"],
            f"{label}.sink arithmetic differs")
    source = result.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict)
            and xlsx.get("implementation") == "source-backed"
            and xlsx.get("cache_mode") == "unmanaged-control"
            and xlsx.get("cache_budget_managed") is False,
            f"{label} normal XLSX source mode differs")
    phases = phase_summary(raw, label)
    unavailable = normal_allocation(raw, label)
    output = result.get("output_sha256")
    require(valid_digest(output), f"{label} output digest is missing")
    source_output = xlsx.get("output_sha256")
    require(isinstance(source_output, list) and source_output == [output],
            f"{label} source output digest differs")
    return {"name": name, "repeat": repeat, "shape": shape,
            "identity": profile_static_identity(result, label),
            "output_sha256": output,
            "phases": phases, "normal_allocation": unavailable}


def profiles() -> dict[str, Any]:
    p = plan()
    binary = binary_identity()
    selected = symbol(binary)
    helper = helper_bindings()
    expected = [
        (f"profile-r{repeat}-{shape}", repeat, shape)
        for repeat in range(1, p["profile"]["repeats"] + 1)
        for shape in p["profile"]["shapes"]
    ]
    actual = {path.name.removesuffix(".receipt.json")
              for path in BASE.glob("*.receipt.json")}
    expected_names = {"build-normal", "symbols"} | {name for name, _, _ in expected}
    require(actual == expected_names,
            f"profile receipt inventory differs: {sorted(actual ^ expected_names)}")
    rows = []
    receipts = [binary["receipt"],
                common(BASE / "symbols.receipt.json", sha(BASE / "source-manifest.json"),
                       binary["binary_sha256"])]
    for name, repeat, shape in expected:
        path = BASE / (name + ".receipt.json")
        receipt = common(path, sha(BASE / "source-manifest.json"), binary["binary_sha256"])
        require(receipt["value"].get("command") ==
                profile_command(name, shape, selected["owner"], binary["binary_path"], p["cpu"]),
                f"{name} command differs")
        entries = receipt["value"].get("artifacts")
        require(isinstance(entries, dict), f"{name} artifacts are missing")
        numbered = []
        for filename in entries:
            if filename.startswith(name + ".callgrind."):
                suffix = filename[len(name + ".callgrind."):]
                if suffix.isdigit():
                    numbered.append(int(suffix))
        require(numbered and sorted(numbered) == list(range(1, max(numbered) + 1)),
                f"{name} numbered Callgrind artifacts are not contiguous")
        required = {name + ".json", name + ".stdout", name + ".stderr",
                    name + ".callgrind"} | {
                        f"{name}.callgrind.{part}" for part in numbered
                    }
        artifacts(receipt["value"], BASE, name, required)
        rows.append(profile_report(BASE / (name + ".json"), name, shape, repeat,
                                   binary, p))
        receipts.append(receipt)
    by_shape = {}
    for row in rows:
        prior = by_shape.setdefault(row["shape"], row["identity"])
        require(row["identity"] == prior,
                f"{row['shape']} output/source identity differs across repeats")
    ordered = sorted(receipts, key=lambda item: item["start"])
    require(all(left["end"] <= right["start"]
                for left, right in zip(ordered, ordered[1:])),
            "build/symbol/profile receipt intervals overlap")
    require(binary["receipt"]["end"] <= receipts[1]["start"],
            "symbols ran before normal build completed")
    require(all(receipts[1]["end"] <= item["start"] for item in receipts[2:]),
            "profile capture ran before symbols completed")
    return {"profile_receipts": len(rows), "serial_receipts": len(receipts),
            "binary_sha256": binary["binary_sha256"], "symbol": selected,
            "helper_bindings": helper,
            "rows": [{key: row[key] for key in
                      ("name", "repeat", "shape", "output_sha256", "phases")}
                     for row in rows]}


def planning_analysis() -> dict[str, Any]:
    profiles()
    need(ANALYZER, "analyze_planning.py")
    recorded = need(HERE / "planning-analysis.json", "planning-analysis.json")
    spec = importlib.util.spec_from_file_location("planning_analyzer_0540", ANALYZER)
    require(spec is not None and spec.loader is not None, "planning analyzer cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    try:
        replayed = module.analyze(False)
    except Exception as error:
        raise EvidenceError(f"planning analyzer replay failed: {error}") from error
    recorded_value = read(recorded)
    require(replayed == recorded_value, "planning analyzer replay differs")
    require(recorded_value.get("status") == "pass"
            and recorded_value.get("plan_sha256") == sha(PLAN)
            and recorded_value.get("owner") in OWNERS
            and isinstance(recorded_value.get("rows"), list)
            and len(recorded_value["rows"]) == 4,
            "planning analysis binding differs")
    return {"report": relative(recorded), "report_sha256": sha(recorded),
            "exact_replay": True, "rows": len(recorded_value["rows"])}


def boundary_analysis() -> dict[str, Any]:
    script_exists, report_exists = BOUNDARY_ANALYZER.exists(), (HERE / "boundary-analysis.json").exists()
    if not script_exists and not report_exists:
        return {"status": "absent", "optional": True}
    need(BOUNDARY_ANALYZER, "analyze_boundary.py")
    recorded = need(HERE / "boundary-analysis.json", "boundary-analysis.json")
    spec = importlib.util.spec_from_file_location("boundary_analyzer_0540", BOUNDARY_ANALYZER)
    require(spec is not None and spec.loader is not None, "boundary analyzer cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    try:
        replayed = module.analyze()
    except Exception as error:
        raise EvidenceError(f"boundary analyzer replay failed: {error}") from error
    value = read(recorded)
    require(replayed == value, "boundary analyzer replay differs")
    require(value.get("status") == "pass" and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 4,
            "boundary analysis binding differs")
    return {"status": "pass", "report": relative(recorded),
            "report_sha256": sha(recorded), "exact_replay": True,
            "rows": len(value["rows"])}


def prior_quality() -> dict[str, Any]:
    value = read(HERE / "quality-reuse.json")
    source_sha = sha(BASE / "source-manifest.json")
    require(value.get("status") == "pass"
            and value.get("source_manifest_sha256") == source_sha
            and value.get("prior_seal_sha256") == sha(PRIOR_SEAL)
            and value.get("successful_test_executions") == 1292,
            "quality reuse binding differs")
    checks = value.get("checks")
    expected = {f"docs/performance/results/change-0539/baseline/{name}.receipt.json"
                for name in QUALITY_NAMES}
    require(isinstance(checks, dict) and set(checks) == expected,
            "quality reuse receipt inventory differs")
    receipts = []
    total_tests = 0
    for name in QUALITY_NAMES:
        key = f"docs/performance/results/change-0539/baseline/{name}.receipt.json"
        item = checks[key]
        path = REPO / key
        require(valid_digest(item.get("sha256")) and sha(path) == item["sha256"],
                f"quality receipt hash differs: {key}")
        receipt = read(path)
        require(receipt.get("exit_code") == 0 and receipt.get("command") == item.get("command"),
                f"quality receipt content differs: {key}")
        start, end = interval(receipt, key)
        artifacts_map = receipt.get("artifacts")
        require(isinstance(artifacts_map, dict)
                and {f"{name}.stdout", f"{name}.stderr"} == set(artifacts_map),
                f"quality receipt artifacts differ: {key}")
        for filename, digest in artifacts_map.items():
            artifact = path.parent / filename
            require(valid_digest(digest) and artifact.is_file() and sha(artifact) == digest,
                    f"quality receipt artifact differs: {key}/{filename}")
        if name == "final-quality-tests":
            total_tests = sum(int(count) for count in re.findall(
                r"test result: ok\. (\d+) passed;",
                (path.parent / f"{name}.stdout").read_text(encoding="utf-8")))
        receipts.append((start, end))
    require(total_tests == 1292, "reused final test count differs")
    ordered = sorted(receipts)
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            "reused quality receipt intervals overlap")
    require(validate_seal(PRIOR_SEAL, PRIOR) > 0, "0539 sealed inventory is invalid")
    return {"checks": len(checks), "successful_test_executions": total_tests,
            "prior_seal_entries": validate_seal(PRIOR_SEAL, PRIOR)}


def cleanup() -> dict[str, Any]:
    value = read(HERE / "cleanup.json")
    require(value.get("removed") == OWNED
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == [],
            "cleanup receipt differs")
    require(value.get("retained_binary_sha256_before_removal", {}).get(
        str(TARGET / "retained-binaries" / "baseline-normal")),
            "cleanup omits retained binary identity")
    require(all(not Path(path).exists() for path in OWNED),
            "owned target remains after cleanup")
    return {"owned_paths_absent": True, "removed": OWNED}


def snapshot() -> dict[str, str]:
    return {relative(path): sha(path) for path in HERE.rglob("*")
            if path.is_file() and not path.is_symlink()}


def all_components() -> dict[str, Any]:
    return {
        "source": validate_source(),
        "environment": validate_environment(),
        "quality": prior_quality(),
        "profiles": profiles(),
        "analysis": planning_analysis(),
        "boundary": boundary_analysis(),
        "cleanup": cleanup(),
        "seal": {"entries": validate_seal(HERE / "SHA256SUMS", HERE)},
    }


def check(name: str) -> Any:
    return {
        "source": validate_source,
        "environment": validate_environment,
        "quality": prior_quality,
        "build": lambda: binary_identity(),
        "symbols": lambda: symbol(binary_identity()),
        "profiles": profiles,
        "analysis": planning_analysis,
        "boundary": boundary_analysis,
        "cleanup": cleanup,
        "seal": lambda: {"entries": validate_seal(HERE / "SHA256SUMS", HERE)},
        "all": all_components,
    }[name]()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=("all", "source", "environment", "quality",
                                                 "build", "symbols", "profiles", "analysis",
                                                 "boundary", "cleanup", "seal"), default="all")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path, help="optional output outside this bundle")
    args = parser.parse_args()
    before = snapshot()
    components = ("source", "environment", "quality", "profiles", "analysis", "boundary",
                  "cleanup", "seal") if args.component == "all" else (args.component,)
    result: dict[str, Any] = {}
    failed = pending = False
    for name in components:
        try:
            result[name] = {"status": "pass", "result": check(name)}
        except Pending as error:
            pending = True
            result[name] = {"status": "pending", "reason": str(error)}
        except EvidenceError as error:
            failed = True
            result[name] = {"status": "fail", "reason": str(error)}
    after = snapshot()
    if before != after:
        failed = True
        result["immutability"] = {"status": "fail", "reason": "verifier mutated bundle"}
    else:
        result["immutability"] = {"status": "pass", "files": len(after)}
    status = "fail" if failed else "incomplete" if pending else "pass"
    output = {"schema": "litchi-0540-attribution-verification-v1", "status": status,
              "performance_claim": "none", "components": result}
    text = json.dumps(output, indent=2, sort_keys=True) + "\n"
    if args.output:
        target = args.output.resolve()
        require(HERE.resolve() not in target.parents and target != HERE.resolve(),
                "refusing output inside evidence bundle")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
    print(text, end="")
    return 1 if failed or (pending and args.strict) else 0


if __name__ == "__main__":
    raise SystemExit(main())
