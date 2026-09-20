#!/usr/bin/env python3
"""Audit the bounded 0704 evidence packet without rebuilding or measuring.

Raw process output is the evidence boundary.  This verifier checks source and
probe bindings, workflow/case counts, sample shape, allocator observables,
profile denominators, focused threshold tests, integration/evidence gates,
the candidate source handoff, the retention observer, and the exact cleanup
scope.  It intentionally permits a PPTX-only source change and never requires
an OOXML shared-codec change.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import statistics
import tempfile
import subprocess
import sys
from pathlib import Path
from typing import Any

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
REPORT = ROOT / "docs" / "performance" / "0704-pptx-bounded-slide-mce-retention.md"
BASE = json.loads((P / "baseline.json").read_text())
CASES = {row["case"]: row["workflow"] for row in json.loads((P / "cases.json").read_text())}
EXPECTED_CASES = {
    "one-real", "one-control", "one-generated", "one-notes-poi", "one-notes-lo",
    "noop-real", "noop-control", "noop-generated", "noop-notes-poi", "noop-notes-lo",
    "two-real", "two-control", "two-generated",
}
EXPECTED_REFUSALS = {
    "small-valid", "generated-12x8-valid", "early-name-error", "late-root-error",
    "late-root-error-mce", "late-missing-relationship", "late-missing-relationship-mce",
    "notes-invalid-tail", "mixed-conformance", "slide-raw-overlimit-16m-to-64m",
}
PHASES = {"baseline", "candidate"}
NATIVE_LEGS = {"a0": "baseline", "a1": "baseline", "a2": "baseline", "a3": "baseline",
               "b0": "candidate", "b1": "candidate"}
NATIVE_COLUMNS = {"capture_ns", "clone_ns", "settext_ns", "commit_ns", "apply_ns", "total_ns"}
ALLOC_PHASES = ("capture", "clone", "settext", "commit", "apply", "total")
ALLOC_FIELDS = ("alloc_calls", "requested_bytes", "baseline_live_bytes", "peak_live_bytes",
                "current_live_bytes", "realloc_calls", "realloc_requested_bytes")
SCRATCH = (
    ROOT.parent / "litchi-target-0704",
    ROOT.parent / "litchi-0704-bin",
    ROOT.parent / "litchi-0704-profile",
    P / "marker-control.pptx",
    ROOT.parent / "litchi-0704-mce-candidate",
)


def fail(message: str) -> None:
    raise AssertionError(message)


def baseline_matches(value: Any) -> bool:
    return isinstance(value, str) and BASE["baseline_head"].startswith(value)


def need(path: Path) -> Path:
    if not path.exists():
        fail(f"missing packet evidence: {path.relative_to(P)}")
    return path


def read(name: str) -> Any:
    return json.loads(need(P / name).read_text())


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_map() -> dict[str, str]:
    return {
        str(path.relative_to(ROOT)): sha(path)
        for owner in ("litchi-pptx", "litchi-ooxml-common", "litchi-opc")
        for path in (ROOT / "crates" / owner).rglob("*.rs")
    }


def file_map(roots: tuple[Path, ...], relative_to: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for root in roots:
        if not root.is_dir():
            fail(f"missing census root: {root}")
        for path in sorted(root.rglob("*")):
            if path.is_file() and not path.is_symlink():
                result[str(path.relative_to(relative_to))] = sha(path)
    return dict(sorted(result.items()))


def observer_production_map() -> dict[str, str]:
    return file_map(
        tuple(ROOT / "crates" / owner for owner in ("litchi-pptx", "litchi-ooxml-common", "litchi-opc")),
        ROOT,
    )


def observer_probe_map() -> dict[str, str]:
    root = P / "retention-probe"
    return {
        name: digest
        for name, digest in file_map((root,), P).items()
        if not name.startswith("retention-probe/results/")
    }


def observer_hash(mapping: dict[str, str]) -> str:
    digest = hashlib.sha256()
    for name, value in mapping.items():
        digest.update(name.encode())
        digest.update(b"\0")
        digest.update(bytes.fromhex(value))
    return digest.hexdigest()


def observer_build_inputs() -> dict[str, str]:
    paths = (
        ROOT / "Cargo.toml",
        ROOT / "Cargo.lock",
        ROOT / "crates" / "litchi-pptx" / "Cargo.toml",
        ROOT / "crates" / "litchi-ooxml-common" / "Cargo.toml",
        ROOT / "crates" / "litchi-opc" / "Cargo.toml",
        P / "retention-probe" / "Cargo.toml",
        P / "retention-probe" / "Cargo.lock",
        P / "source-census-candidate.json",
    )
    return {str(path.relative_to(ROOT)): sha(path) for path in paths}


def verify_json_hash(path: Path, key: str, value: str) -> None:
    if sha(path) != value:
        fail(f"hash mismatch for {path.relative_to(P)} ({key})")


def verify_binary(path_value: str, expected: str, label: str) -> None:
    path = Path(path_value)
    if path.exists():
        verify_json_hash(path, label, expected)
        return
    cleanup = P / "cleanup.json"
    if not cleanup.exists():
        fail(f"missing {label} binary: {path}")
    rows = {row["path"]: row for row in read("cleanup.json").get("binaries", [])}
    row = rows.get(str(path))
    if row is None or row.get("binary_sha256") != expected:
        fail(f"removed {label} binary lacks a matching cleanup receipt: {path}")


def verify_probe_inputs() -> None:
    for name in ("probe", "refusal-probe"):
        lock = P / name / "Cargo.lock"
        manifest = P / name / "Cargo.toml"
        need(lock)
        need(manifest)
        result = subprocess.run(
            ["cargo", "metadata", "--manifest-path", str(manifest), "--locked", "--no-deps"],
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )
        if result.returncode:
            fail(f"{name} lockfile does not validate with cargo metadata: {result.stderr[-500:]}")
    if 'name = "probe0704"' not in (P / "probe/Cargo.lock").read_text():
        fail("native standalone lock is not named probe0704")
    if 'name = "probe0704-refusal"' not in (P / "refusal-probe/Cargo.lock").read_text():
        fail("refusal standalone lock is not named probe0704-refusal")
    probe_files = {
        str(path.relative_to(P)): sha(path)
        for path in (P / "probe").rglob("*")
        if path.is_file()
    }
    refusal_files = {
        str(path.relative_to(P)): sha(path)
        for path in (P / "refusal-probe").rglob("*")
        if path.is_file()
    }
    for phase in PHASES:
        for row in read(f"builds-{phase}.json"):
            if row["probe_sha256"] != probe_files:
                fail(f"{phase} native build is not bound to the frozen probe")
        refusal = read(f"build-refusal-{phase}.json")
        if refusal["probe_sha256"] != refusal_files:
            fail(f"{phase} refusal build is not bound to the frozen probe")


OBSERVER_CHARGE_FIELDS = {
    "source_retained_mce_bytes",
    "clone_retained_mce_bytes",
    "clone_after_release_bytes",
    "source_after_clone_release_bytes",
    "transaction_retained_mce_bytes",
    "transaction_after_release_bytes",
    "transaction_after_first_edit_bytes",
    "transaction_after_second_edit_bytes",
    "commit_retained_mce_bytes",
    "commit_snapshot_retained_mce_bytes",
    "release_commit_retained_mce_bytes",
    "commit_after_release_bytes",
    "commit_reference_retained_mce_bytes",
    "published_retained_mce_bytes",
}


def packet_path(relative_name: str) -> Path:
    path = (P / relative_name).resolve()
    try:
        path.relative_to(P.resolve())
    except ValueError:
        fail(f"observer receipt escapes the packet: {relative_name}")
    return need(path)


def parse_observer_output(path: Path, source: str, repeats: int) -> None:
    lines = path.read_text().splitlines()
    required = {
        "probe\t0704-retention-observer",
        f"source\t{source}",
        f"repeats\t{repeats}",
        "parity\tdefault_vs_off\ttrue",
        "parity\tdefault_vs_tiny\ttrue",
        "run_complete\ttrue",
    }
    if not required.issubset(lines):
        fail(f"observer output omitted required lines: {path}")
    starts = [index for index, line in enumerate(lines) if line.startswith("record\t")]
    if len(starts) != repeats * 3:
        fail(f"observer output record count mismatch: {path}")
    records: list[dict[str, str]] = []
    for offset, start in enumerate(starts):
        stop = starts[offset + 1] if offset + 1 < len(starts) else len(lines)
        fields = lines[start].split("\t")
        if len(fields) != 4:
            fail(f"observer record header is malformed: {path}")
        row: dict[str, str] = {
            "source": fields[1], "budget": fields[2], "repeat": fields[3]
        }
        for line in lines[start + 1 : stop]:
            parts = line.split("\t", 1)
            if len(parts) == 2:
                row[parts[0]] = parts[1]
        missing = {
            "budget_bytes",
            "before_revision",
            "commit_revision",
            "after_revision",
            "patch_bytes",
            "patch_changed",
            "first_changed",
            "second_changed",
        } | OBSERVER_CHARGE_FIELDS
        if missing - row.keys():
            fail(f"observer record omitted fields {sorted(missing - row.keys())}: {path}")
        try:
            repeat = int(row["repeat"])
            budget_bytes = int(row["budget_bytes"])
            patch_bytes = int(row["patch_bytes"])
            charges = {name: int(row[name]) for name in OBSERVER_CHARGE_FIELDS}
        except (KeyError, ValueError):
            fail(f"observer record has non-numeric charge fields: {path}")
        if repeat not in range(repeats) or patch_bytes < 0 or any(value < 0 for value in charges.values()):
            fail(f"observer record has invalid repeat or negative observation: {path}")
        expected_budget = {"default": 1024 * 1024, "off": 0, "tiny": 1}.get(row["budget"])
        if expected_budget is None or budget_bytes != expected_budget:
            fail(f"observer budget census mismatch: {path}")
        if charges["source_retained_mce_bytes"] > budget_bytes:
            fail(f"observer source retention exceeded its declared budget: {path}")
        if row["budget"] == "off" and charges["source_retained_mce_bytes"] != 0:
            fail(f"observer off policy retained bytes: {path}")
        records.append(row)
    keys = {(row["repeat"], row["budget"]) for row in records}
    if len(keys) != repeats * 3 or {row["budget"] for row in records} != {"default", "off", "tiny"}:
        fail(f"observer repeat/budget census mismatch: {path}")
    if source == "generated:12x8":
        if any(int(row["source_retained_mce_bytes"]) or int(row["published_retained_mce_bytes"])
               for row in records):
            fail(f"marker-free observer fixture retained MCE bytes: {path}")
        expected_target = "target1\tslide:0\tshape:0"
    else:
        if any(int(row["source_retained_mce_bytes"]) <= 0
               for row in records if row["budget"] == "default"):
            fail(f"marker-bearing observer fixture retained no default MCE bytes: {path}")
        expected_target = "target1\tslide:1\tshape:1"
    if expected_target not in lines:
        fail(f"observer target receipt missing: {path}")
    if any(line.startswith(("capture_ns\t", "seconds\t", "rss\t", "alloc_")) for line in lines):
        fail(f"observer output contains an unrequested performance measurement: {path}")


def verify_observer() -> None:
    for name in (
        "build-retention.py",
        "run-retention.py",
        "retention-observer-README.md",
        "retention-probe/Cargo.toml",
        "retention-probe/Cargo.lock",
        "retention-probe/src/main.rs",
    ):
        need(P / name)
    if 'name = "probe0704-retention"' not in (P / "retention-probe/Cargo.lock").read_text():
        fail("retention observer standalone lock has the wrong package name")
    build = read("retention-build.json")
    if build.get("probe") != "0704-retention-observer":
        fail("retention observer build identity mismatch")
    production = observer_production_map()
    probe = observer_probe_map()
    if build.get("production_source_sha256") != production:
        fail("retention observer production source census mismatch")
    if build.get("production_file_count") != len(production):
        fail("retention observer production file count mismatch")
    if build.get("probe_sha256") != probe or build.get("probe_file_count") != len(probe):
        fail("retention observer probe census mismatch")
    if build.get("build_inputs_sha256") != observer_build_inputs():
        fail("retention observer build input census mismatch")
    if build.get("production_source_hash") != observer_hash(production):
        fail("retention observer production source hash mismatch")
    if build.get("probe_source_hash") != observer_hash(probe):
        fail("retention observer probe source hash mismatch")
    combined = dict(production)
    combined.update(probe)
    if build.get("source_hash") != observer_hash(combined):
        fail("retention observer combined source hash mismatch")
    rust = {name: digest for name, digest in production.items() if name.endswith(".rs")}
    other = {name: digest for name, digest in production.items() if not name.endswith(".rs")}
    if rust != read("source-census-candidate.json")["source_sha256"]:
        fail("observer Rust census does not match candidate")
    for prefix, mapping in (("production_rust_source", rust), ("production_other", other)):
        if build.get(prefix + "_sha256") != mapping or build.get(prefix + "_hash") != observer_hash(mapping):
            fail("observer separate Rust/other census mismatch")
    verify_binary(build["binary"], build["binary_sha256"], "retention observer")

    runs = read("retention-runs.json")
    if runs.get("probe") != "0704-retention-observer" or runs.get("repeats") != 2:
        fail("retention observer run census mismatch")
    for key in ("binary", "binary_sha256", "production_source_hash", "probe_source_hash", "source_hash"):
        if runs.get(key) != build.get(key):
            fail(f"retention observer run/build binding mismatch: {key}")
    if runs.get("performance_claim") != "none":
        fail("retention observer made an unintended performance claim")
    cases = runs.get("cases", [])
    if {row.get("label") for row in cases} != {"real", "generated"} or len(cases) != 2:
        fail("retention observer fixture census mismatch")
    real = ROOT / "test-data" / "libreoffice-core" / "sd" / "qa" / "unit" / "data" / "pptx" / "slide-section-test.pptx"
    if not real.is_file():
        fail(f"retention observer real fixture is missing: {real}")
    for row in cases:
        if row.get("exit_code") != 0 or row.get("binary") != build.get("binary"):
            fail(f"invalid retention observer run receipt: {row}")
        if row.get("binary_sha256") != build.get("binary_sha256"):
            fail("retention observer run binary hash mismatch")
        output = packet_path(row["output"])
        stderr = packet_path(row["stderr"])
        if sha(output) != row.get("output_sha256") or sha(stderr) != row.get("stderr_sha256"):
            fail(f"retention observer raw hash mismatch: {output}")
        if row.get("label") == "real":
            if row.get("source") != str(real) or row.get("source_sha256") != sha(real):
                fail("retention observer real fixture binding mismatch")
        elif row.get("source") != "generated:12x8" or row.get("source_sha256") is not None:
            fail("retention observer generated fixture binding mismatch")
        parse_observer_output(output, row["source"], runs["repeats"])


MECHANISM_BUILD_NAMES = (
    "mechanism/build.json",
    "mechanism/mechanism-build.json",
    "mechanism/build-mechanism.json",
)
MECHANISM_RUN_NAMES = (
    "mechanism/runs.json",
    "mechanism/mechanism-runs.json",
    "mechanism/trace-runs.json",
    "mechanism/runs-candidate.json",
)
MECHANISM_CODEC = "crates/litchi-ooxml-common/src/mce/codec.rs"


def mechanism_file(relative_name: str) -> Path:
    candidates = [P / relative_name]
    if not relative_name.startswith("mechanism/"):
        candidates.append(P / "mechanism" / relative_name)
    for path in candidates:
        if path.is_file():
            return path
    fail(f"missing mechanism evidence: {relative_name}")


def mechanism_build_receipt() -> tuple[Path, dict[str, Any]] | None:
    paths = [P / name for name in MECHANISM_BUILD_NAMES if (P / name).is_file()]
    if not paths:
        return None
    if len(paths) != 1:
        fail("multiple mechanism build receipts are present")
    path = paths[0]
    return path, json.loads(path.read_text())


def mechanism_digest_path(value: str) -> Path:
    candidates = [Path(value)] if Path(value).is_absolute() else [P / "mechanism" / value, P / value]
    for path in candidates:
        if path.is_file():
            return path
    fail(f"missing mechanism artifact: {value}")


def mechanism_source_map(build: dict[str, Any], *names: str) -> dict[str, str] | None:
    for name in names:
        value = build.get(name)
        if isinstance(value, dict):
            return value
    return None


def parse_mechanism_trace(path: Path) -> int:
    prefix = "LITCHI0704_MCE "
    boundary = "LITCHI0704_BOUNDARY "
    required = {
        "call", "phase", "profile", "raw_ptr", "raw_len", "raw_sha256",
        "output_ptr", "output_len", "output_capacity", "ownership",
        "output_sha256", "limits_input", "limits_output", "limits_depth",
        "limits_bindings", "limits_directives", "limits_choices", "status",
    }
    calls: list[int] = []
    for line in path.read_text().splitlines():
        line = line.strip()
        if line.startswith(boundary):
            values = dict(token.split("=", 1) for token in line[len(boundary):].split() if "=" in token)
            if values.get("event") != "begin" or not values.get("phase"):
                fail(f"malformed mechanism boundary in {path}")
            continue
        if not line.startswith(prefix):
            continue
        values = dict(token.split("=", 1) for token in line[len(prefix):].split() if "=" in token)
        if set(values) != required:
            fail(f"mechanism trace field census mismatch: {path}")
        try:
            numeric = {key: int(values[key], 0) if key.endswith("ptr") else int(values[key])
                       for key in ("call", "raw_ptr", "raw_len", "output_ptr", "output_len", "output_capacity")}
        except (KeyError, ValueError):
            fail(f"mechanism trace numeric field mismatch: {path}")
        if any(value < 0 for value in numeric.values()) or values["profile"] != "default-ooxml":
            fail(f"mechanism trace contains an invalid observation: {path}")
        if values["status"] not in {"ok", "error"}:
            fail(f"mechanism trace contains an invalid status: {path}")
        if values["status"] == "error" and values["ownership"] != "error":
            fail(f"mechanism error trace has the wrong ownership: {path}")
        if values["status"] == "ok" and values["ownership"] not in {"borrowed", "owned"}:
            fail(f"mechanism success trace has the wrong ownership: {path}")
        calls.append(numeric["call"])
    if calls != list(range(len(calls))):
        fail(f"mechanism trace call sequence is not contiguous: {path}")
    if not calls:
        fail(f"mechanism trace contains no direct MCE call observations: {path}")
    return len(calls)


def verify_mechanism() -> None:
    if not (P / "mechanism").exists():
        return
    found = mechanism_build_receipt()
    if found is None:
        fail("mechanism directory has no build receipt")
    build_path, build = found
    if build.get("status") not in (None, "completed"):
        fail("mechanism build did not complete")
    if build.get("diagnostic_only") is not True or build.get("performance_claim") != "none":
        fail("mechanism lane is not marked diagnostic-only")
    if not baseline_matches(build.get("baseline_head")):
        fail("mechanism build baseline mismatch")
    current = source_map()
    before = mechanism_source_map(
        build,
        "candidate_source_sha256",
        "pre_instrumentation_source_sha256",
        "original_source_sha256",
    )
    restored = mechanism_source_map(
        build,
        "postrestore_candidate_source_sha256",
        "postrestore_source_sha256",
        "restored_source_sha256",
    )
    if before is None or restored is None:
        fail("mechanism receipt lacks pre-instrumentation or post-restore candidate source hashes")
    if before != current:
        fail("mechanism pre-instrumentation source map differs from the candidate")
    if restored != current:
        fail("mechanism post-restore source map differs from the candidate")
    original_codec = build.get("original_codec_sha256")
    if original_codec != BASE["source_sha256"].get(MECHANISM_CODEC):
        fail("mechanism original codec hash is not the baseline hash")
    restored_codec = build.get("restored_codec_sha256") or build.get("postrestore_codec_sha256")
    if restored_codec != current.get(MECHANISM_CODEC):
        fail("mechanism restored codec hash differs from the candidate")
    instrumented = mechanism_source_map(build, "instrumented_source_sha256")
    instrumented_codec = build.get("instrumented_codec_sha256")
    if instrumented is not None:
        if set(instrumented) != set(current) or instrumented.get(MECHANISM_CODEC) != instrumented_codec:
            fail("mechanism instrumented source map is incomplete")
        changed = [name for name in current if instrumented[name] != current[name]]
        if changed != [MECHANISM_CODEC] or instrumented_codec == current.get(MECHANISM_CODEC):
            fail("mechanism instrumentation changed a source outside the codec")
    else:
        fail("mechanism receipt lacks an instrumented source map")
    patch_ref = build.get("patch") or build.get("instrumentation_patch")
    if patch_ref is None:
        patches = sorted((P / "mechanism").glob("*.patch"))
        if len(patches) != 1:
            fail("mechanism receipt does not identify a unique instrumentation patch")
        patch = patches[0]
    else:
        patch = mechanism_digest_path(patch_ref)
    patch_hash = build.get("patch_sha256") or build.get("instrumentation_patch_sha256")
    if not patch_hash or sha(patch) != patch_hash:
        fail("mechanism instrumentation patch hash mismatch")
    patch_lines = patch.read_text(errors="strict").splitlines()
    patch_paths: set[str] = set()
    for index, line in enumerate(patch_lines):
        if line.startswith("diff --git "):
            fields = line.split()
            if len(fields) != 4 or not fields[3].startswith("b/"):
                fail("mechanism instrumentation patch has a malformed git header")
            patch_paths.add(fields[3][2:])
        elif line.startswith("--- a/"):
            if index + 1 >= len(patch_lines) or not patch_lines[index + 1].startswith("+++ b/"):
                fail("mechanism instrumentation patch has a malformed unified header")
            patch_paths.add(patch_lines[index + 1][6:])
    if patch_paths != {MECHANISM_CODEC}:
        fail("mechanism instrumentation patch changes more than the shared codec")
    probe_files = build.get("probe_sha256")
    if not isinstance(probe_files, dict) or not probe_files:
        fail("mechanism receipt has no probe source census")
    for name, digest in probe_files.items():
        path = mechanism_digest_path(name)
        if sha(path) != digest:
            fail(f"mechanism probe source hash mismatch: {name}")
    for step in build.get("steps", []):
        if step.get("exit_code") != 0:
            fail(f"mechanism build step failed: {step.get('name')}")
        log_name = step.get("log") or f"{step.get('name')}.log"
        log = mechanism_digest_path(log_name)
        if step.get("log_sha256") and sha(log) != step["log_sha256"]:
            fail(f"mechanism build log hash mismatch: {log}")
    if build.get("workspace_lock_sha256") and build["workspace_lock_sha256"] != sha(ROOT / "Cargo.lock"):
        fail("mechanism build changed the workspace Cargo.lock")
    verify_binary(build["binary"], build["binary_sha256"], "mechanism")

    run_paths = [P / name for name in MECHANISM_RUN_NAMES if (P / name).is_file()]
    if len(run_paths) != 1:
        fail("mechanism lane must retain exactly one run receipt")
    run_data = json.loads(run_paths[0].read_text())
    rows = run_data if isinstance(run_data, list) else (
        run_data.get("runs") or run_data.get("records") or run_data.get("cases")
    )
    if not isinstance(rows, list) or len(rows) != 12:
        fail("mechanism lane must retain twelve fresh trace runs")
    run_keys: set[tuple[str, str, int]] = set()
    for row in rows:
        if row.get("exit_code") != 0 or row.get("binary_sha256") != build["binary_sha256"]:
            fail("invalid mechanism trace run receipt")
        key = (row.get("case"), row.get("workflow"), row.get("repeat"))
        if key in run_keys or key[0] not in {"real", "generated"} or key[1] not in {"noop", "one", "two"} or key[2] not in {0, 1}:
            fail("mechanism trace run census has a duplicate or invalid key")
        run_keys.add(key)
        if row.get("candidate_source_map_sha256") != build.get("candidate_source_map_sha256"):
            fail("mechanism trace run source-map binding mismatch")
        if not row.get("stderr"):
            fail("mechanism trace run has no raw stderr trace")
        trace = mechanism_digest_path(row["stderr"])
        trace_count = parse_mechanism_trace(trace)
        if row.get("trace_call_count") is not None and row["trace_call_count"] != trace_count:
            fail(f"mechanism trace call count mismatch: {trace}")
        for stream, digest_key in (("stdout", "stdout_sha256"), ("stderr", "stderr_sha256")):
            if row.get(stream):
                artifact = mechanism_digest_path(row[stream])
                if row.get(digest_key) and sha(artifact) != row[digest_key]:
                    fail(f"mechanism {stream} hash mismatch: {artifact}")
        if row.get("case") == "real" and row.get("source_archive_sha256"):
            real = ROOT / "test-data" / "libreoffice-core" / "sd" / "qa" / "unit" / "data" / "pptx" / "slide-section-test.pptx"
            if row["source_archive_sha256"] != sha(real):
                fail("mechanism real fixture hash mismatch")
        if any(key in row for key in ("timing", "timings", "elapsed_ns", "rss_bytes", "alloc_calls")):
            fail("mechanism trace receipt contains a performance measurement")
    if run_keys != {(case, workflow, repeat) for case in ("real", "generated")
                    for workflow in ("noop", "one", "two") for repeat in (0, 1)}:
        fail("mechanism trace run coverage mismatch")
    summary_path = P / "mechanism" / "trace-summary.json"
    if not summary_path.is_file():
        fail("mechanism trace summary is missing")
    summary = json.loads(summary_path.read_text())
    if summary.get("diagnostic_only") is not True or summary.get("timing_evidence") is not False:
        fail("mechanism trace summary is not diagnostic-only")
    summary_runs = summary.get("runs")
    if not isinstance(summary_runs, list) or len(summary_runs) != 12 or any(row.get("calls", 0) <= 0 for row in summary_runs):
        fail("mechanism trace summary lacks direct call counts for all runs")


def verify_sources() -> None:
    baseline = read("source-census-baseline.json")
    if baseline["source_file_count"] != 602 or baseline["source_sha256"] != BASE["source_sha256"]:
        fail("baseline source census is not the recorded 602-file map")
    current = source_map()
    candidate = read("source-census-candidate.json")
    if candidate["source_sha256"] != current:
        fail("candidate source census does not match the current checkout")
    if candidate["source_file_count"] < 603:
        fail(f"candidate source census did not add a PPTX source: {candidate['source_file_count']}")
    if candidate["removed_paths"]:
        fail(f"candidate removed sources: {candidate['removed_paths'][:4]}")
    outside = [
        name for name in candidate["added_paths"] + candidate["changed_paths"]
        if not name.startswith("crates/litchi-pptx/")
    ]
    if outside:
        fail(f"candidate source change outside PPTX: {outside[:5]}")
    for name, digest in BASE["source_sha256"].items():
        if name.startswith(("crates/litchi-ooxml-common/", "crates/litchi-opc/")) and current[name] != digest:
            fail(f"shared source changed in PPTX packet: {name}")
    if not any(re.search(r"memo|projection|retention|budget|cache", name, re.I)
               for name in candidate["added_paths"] + candidate["changed_paths"]):
        fail("candidate source census does not identify a memo/retention source path")


def verify_report_binding() -> None:
    if not REPORT.is_file():
        fail(f"missing 0704 report at the required path: {REPORT}")
    text = REPORT.read_text()
    for marker in ("d48523eec2", "performance_claim: none", "results/change-0704/README.md"):
        if marker not in text:
            fail(f"0704 report is missing required binding: {marker}")


def witness_packet_path(value: str) -> Path:
    path = (P / value).resolve()
    try:
        path.relative_to(P.resolve())
    except ValueError:
        fail(f"candidate witness path escapes the packet: {value}")
    return path


def historical_handoff(witness: dict[str, Any]) -> dict[str, Any]:
    handoff = witness.get("historical_handoff")
    if handoff is not None:
        return handoff
    saved = P / "preflight" / "candidate-coder-handoff-witness.json"
    if saved.exists():
        return json.loads(saved.read_text())
    return witness


def verify_final_candidate_witness(witness: dict[str, Any]) -> None:
    required = {
        "final_candidate_patch",
        "final_candidate_patch_sha256",
        "final_candidate_source_sha256",
        "final_candidate_patch_files",
    }
    if not required.issubset(witness):
        fail("final candidate source-diff witness has not been recorded")
    patch = need(witness_packet_path(witness["final_candidate_patch"]))
    if sha(patch) != witness["final_candidate_patch_sha256"]:
        fail("final candidate source-diff hash mismatch")
    source_hashes = witness["final_candidate_source_sha256"]
    patch_files = witness["final_candidate_patch_files"]
    if not source_hashes or set(source_hashes) != set(patch_files):
        fail("final candidate source-diff file census is incomplete")
    if any(not name.startswith("crates/litchi-pptx/") for name in source_hashes):
        fail("final candidate source-diff escapes the PPTX source boundary")
    candidate = read("source-census-candidate.json")
    expected = set(candidate["added_paths"]) | set(candidate["changed_paths"])
    if set(source_hashes) != expected:
        fail("final candidate source-diff does not cover every candidate production/test source")
    for name, digest in source_hashes.items():
        path = ROOT / name
        if not path.is_file() or sha(path) != digest:
            fail(f"final candidate source-diff source mismatch: {path}")
    diff_paths: set[str] = set()
    for line in patch.read_text(errors="strict").splitlines():
        if not line.startswith("diff --git "):
            continue
        fields = line.split()
        if len(fields) != 4 or not fields[2].startswith("a/") or not fields[3].startswith("b/"):
            fail("final candidate source-diff has a malformed file header")
        left, right = fields[2][2:], fields[3][2:]
        if left != right:
            fail("final candidate source-diff contains a rename")
        diff_paths.add(right)
    if diff_paths != set(patch_files):
        fail("final candidate source-diff headers do not match its witness census")
    with tempfile.TemporaryDirectory(prefix="litchi-0704-source-audit-") as directory:
        root = Path(directory)
        for name in source_hashes:
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            if name in BASE["source_sha256"]:
                content = subprocess.check_output(["git", "show", BASE["baseline_head"] + ":" + name], cwd=ROOT)
                if hashlib.sha256(content).hexdigest() != BASE["source_sha256"][name]:
                    fail("baseline git content mismatch")
                target.write_bytes(content)
        subprocess.run(["git", "apply", str(patch)], cwd=root, check=True)
        if {name: sha(root / name) for name in source_hashes} != source_hashes:
            fail("candidate patch replay does not reconstruct final sources")
    if any(not name.startswith("crates/litchi-pptx/")
           for name in witness.get("policy_files_excluded", [])):
        fail("candidate policy exclusion witness escapes the PPTX source boundary")


def verify_historical_handoff(witness: dict[str, Any]) -> None:
    handoff = historical_handoff(witness)
    if not baseline_matches(handoff.get("baseline")):
        fail("historical coder handoff baseline mismatch")
    patch_ref = handoff.get("candidate_patch", "candidate-production.patch")
    patch = need(witness_packet_path(patch_ref))
    if sha(patch) != handoff.get("candidate_patch_sha256"):
        fail("historical coder handoff patch hash mismatch")
    source_hashes = handoff.get("candidate_source_sha256", {})
    patch_files = handoff.get("candidate_patch_files", [])
    if not source_hashes or set(source_hashes) != set(patch_files):
        fail("historical coder handoff source census is incomplete")
    if any(not name.startswith("crates/litchi-pptx/") for name in source_hashes):
        fail("historical coder handoff escapes the PPTX source boundary")
    if handoff.get("cargo_or_performance_run") is not False:
        fail("historical coder handoff does not state that no Cargo/performance run occurred")
    worktree_value = handoff.get("candidate_worktree")
    if not worktree_value:
        fail("historical coder handoff has no worktree path")
    worktree = Path(worktree_value)
    if worktree.exists():
        if not worktree.is_dir() or worktree.is_symlink():
            fail("historical coder handoff worktree is not a real directory")
        for name, digest in source_hashes.items():
            path = worktree / name
            if not path.is_file() or sha(path) != digest:
                fail(f"historical coder handoff source mismatch: {path}")
    else:
        cleanup = P / "cleanup.json"
        if not cleanup.exists():
            fail("historical coder worktree disappeared without cleanup receipt")
        retained = read("cleanup.json").get("historical_worktree_witness", {})
        if retained.get("path") != str(worktree):
            fail("cleanup receipt lost the historical coder worktree path")
        if retained.get("candidate_patch_sha256") != handoff.get("candidate_patch_sha256"):
            fail("cleanup receipt lost the historical coder patch witness")
        if retained.get("candidate_source_sha256") != source_hashes:
            fail("cleanup receipt lost the historical coder source witness")


def verify_candidate_witness() -> None:
    witness = read("candidate-source-witness.json")
    if not baseline_matches(witness.get("baseline")):
        fail("candidate source witness baseline mismatch")
    verify_final_candidate_witness(witness)
    verify_historical_handoff(witness)


def verify_scripts_manifest() -> None:
    manifest = read("scripts-manifest.json")
    if manifest.get("baseline_head") != BASE["baseline_head"]:
        fail("scripts manifest baseline head mismatch")
    if manifest.get("baseline_source_file_count") != 602:
        fail("scripts manifest baseline source count mismatch")
    entries = manifest.get("packet_files", {})
    if not entries:
        fail("scripts manifest is empty")
    required = {
        "build-retention.py",
        "run-retention.py",
        "retention-observer-README.md",
        "retention-probe/Cargo.toml",
        "retention-probe/Cargo.lock",
        "retention-probe/src/main.rs",
    }
    if not required.issubset(entries):
        fail(f"scripts manifest omitted retention observer inputs: {sorted(required - set(entries))}")
    if (P / "mechanism").exists() and not any(name.startswith("mechanism/") for name in entries):
        fail("scripts manifest omitted the separate mechanism lane")
    for name, record in entries.items():
        path = P / name
        if not path.exists() or sha(path) != record.get("sha256") or path.stat().st_size != record.get("bytes"):
            fail(f"scripts manifest hash mismatch: {name}")


def verify_builds() -> None:
    expected_probe = {
        "native", "allocations"
    }
    for phase in PHASES:
        rows = read(f"builds-{phase}.json")
        if {row.get("label") for row in rows} != expected_probe or len(rows) != 2:
            fail(f"{phase} native/allocation build census mismatch")
        for row in rows:
            if row.get("phase") != phase or row.get("exit_code") != 0 or row.get("rustflags") != "-D warnings":
                fail(f"invalid {phase} build receipt: {row}")
            verify_binary(row["binary"], row["binary_sha256"], f"{phase} {row['label']}")
            if row["label"] == "native" and not row["binary"].endswith(f"{phase}-native"):
                fail(f"native binary name is not frozen for {phase}")
        refusal = read(f"build-refusal-{phase}.json")
        if refusal.get("phase") != phase or refusal.get("exit_code") != 0:
            fail(f"invalid {phase} refusal build receipt")
        verify_binary(refusal["binary"], refusal["binary_sha256"], f"{phase} refusal")


def parse_native(path: Path) -> tuple[dict[str, list[str]], list[dict[str, int]]]:
    lines = need(path).read_text().splitlines()
    metadata: dict[str, list[str]] = {}
    rows: list[dict[str, int]] = []
    header: list[str] | None = None
    for line in lines:
        fields = line.split("\t")
        if fields[0] == "sample":
            header = fields
        elif fields[0].isdigit():
            if header is None or len(fields) != len(header):
                fail(f"{path}: malformed native row")
            values = [int(value) for value in fields]
            if any(value < 0 for value in values):
                fail(f"{path}: negative native sample")
            rows.append(dict(zip(header, values)))
        elif fields[0]:
            metadata[fields[0]] = fields[1:]
    if header is None or len(rows) != 100 or [row["sample"] for row in rows] != list(range(100)):
        fail(f"{path}: expected exactly 100 contiguous native samples")
    if metadata.get("probe") != ["0704"] or metadata.get("samples") != ["100"] or metadata.get("warmups") != ["5"]:
        fail(f"{path}: native provenance mismatch")
    if not NATIVE_COLUMNS.issubset(set(header[1:])):
        fail(f"{path}: native phase columns missing")
    return metadata, rows


def verify_native() -> None:
    seen: set[tuple[str, str]] = set()
    for filename, expected_legs in (("native-runs-baseline.json", {"a0", "a1"}),
                                    ("native-runs-compare.json", {"a2", "a3", "b0", "b1"})):
        records = read(filename)
        if len(records) != len(CASES) * len(expected_legs):
            fail(f"{filename}: expected {len(CASES) * len(expected_legs)} records")
        for record in records:
            case = record.get("case")
            leg = record.get("leg")
            if case not in CASES or leg not in expected_legs or record.get("exit_code") != 0:
                fail(f"invalid native run record: {record}")
            key = (case, leg)
            if key in seen:
                fail(f"duplicate native run record: {key}")
            seen.add(key)
            output = P / record["output"]
            if sha(output) != record["output_sha256"]:
                fail(f"native output hash mismatch: {output}")
            metadata, rows = parse_native(output)
            workflow = CASES[case]
            if metadata.get("workflow", ["one"]) != [workflow]:
                fail(f"{output}: workflow metadata mismatch")
            if record.get("phase") != NATIVE_LEGS[leg]:
                fail(f"{output}: phase/leg mismatch")
            if workflow == "one":
                required = {"before_revision_sha256", "after_revision_sha256", "candidate_archive_sha256",
                            "reopened_target_text_sha256", "correctness_target_text"}
            elif workflow == "noop":
                required = {"before_revision_sha256", "after_revision_sha256", "before_archive_sha256",
                            "after_archive_sha256", "commit_is_changed", "revision_identical",
                            "output_identical", "correctness_target_text_sha256"}
            else:
                required = {"before_revision_sha256", "after_revision_sha256", "candidate_archive_sha256",
                            "reopened_target1_text_sha256", "reopened_target2_text_sha256",
                            "correctness_target_text"}
            if not required.issubset(metadata):
                fail(f"{output}: semantic validation metadata missing {sorted(required - set(metadata))}")
            if any(set(row) < NATIVE_COLUMNS | {"sample"} for row in rows):
                fail(f"{output}: phase sample columns are incomplete")
    if seen != {(case, leg) for case in CASES for leg in NATIVE_LEGS}:
        fail("native case/leg coverage mismatch")
    if read("native-summary.json") is None or read("native-comparisons.json") is None:
        fail("native summaries are missing")


def parse_refusal(path: Path) -> None:
    lines = need(path).read_text().splitlines()
    header: dict[str, str] = {}
    cases: dict[str, list[int]] = {}
    current: str | None = None
    for line in lines:
        fields = line.split("\t")
        if fields[0] == "case":
            current = fields[1]
            if current in cases:
                fail(f"{path}: duplicate refusal case")
            cases[current] = []
        elif fields[0] == "sample_ns":
            continue
        elif fields[0].isdigit():
            if current is None or int(fields[0]) != len(cases[current]):
                fail(f"{path}: refusal sample index mismatch")
            cases[current].append(int(fields[1]))
        elif fields[0] == "all_iterations_passed":
            header[fields[0]] = fields[1]
        elif current is None:
            header[fields[0]] = fields[1]
        else:
            # Setup metadata is intentionally retained but not interpreted here.
            pass
    expected_header = {
        "probe": "0693-refusal",
        "samples": "100",
        "warmups": "5",
        "all_iterations_passed": "true",
    }
    if any(header.get(key) != value for key, value in expected_header.items()):
        fail(f"{path}: refusal header mismatch")
    if set(cases) != EXPECTED_REFUSALS or any(len(values) != 100 for values in cases.values()):
        fail(f"{path}: refusal case/sample census mismatch")


def verify_refusal() -> None:
    seen: set[str] = set()
    for filename, expected_legs in (("refusal-runs-baseline.json", {"a0", "a1"}),
                                    ("refusal-runs-compare.json", {"a2", "a3", "b0", "b1"})):
        records = read(filename)
        if len(records) != len(expected_legs):
            fail(f"{filename}: refusal leg count mismatch")
        for record in records:
            if record.get("leg") not in expected_legs or record.get("exit_code") != 0:
                fail(f"invalid refusal run record: {record}")
            if record["leg"] in seen:
                fail(f"duplicate refusal leg: {record['leg']}")
            seen.add(record["leg"])
            output = P / record["output"]
            if sha(output) != record["output_sha256"] or sha(output.with_suffix(".stderr")) != record["stderr_sha256"]:
                fail(f"refusal raw hash mismatch: {output}")
            parse_refusal(output)
    if seen != set(NATIVE_LEGS):
        fail("refusal leg coverage mismatch")
    for name in ("refusal-summary.json", "refusal-comparisons.json", "refusal-bindings.json"):
        need(P / name)


def parse_allocation(path: Path) -> None:
    lines = need(path).read_text().splitlines()
    header = next((line.split("\t") for line in lines if line.startswith("sample\t")), None)
    if header is None:
        fail(f"{path}: allocation header missing")
    rows = [line.split("\t") for line in lines if line and line.split("\t")[0].isdigit()]
    declared = next((int(line.split("\t")[1]) for line in lines if line.startswith("samples\t")), None)
    if declared != 3 or len(rows) != 3:
        fail(f"{path}: allocation sample count mismatch")
    columns = set(header[1:])
    required = {f"{phase}_{field}" for phase in ALLOC_PHASES for field in ALLOC_FIELDS}
    if not required.issubset(columns):
        fail(f"{path}: allocation columns missing")


def verify_allocations() -> None:
    baseline = read("allocation-runs-baseline.json")
    candidate = read("allocation-runs-candidate.json")
    if len(baseline) != len(CASES) or len(candidate) != len(CASES):
        fail("allocation run census mismatch")
    for phase, records in (("baseline", baseline), ("candidate", candidate)):
        for record in records:
            if record.get("phase") != phase or record.get("exit_code") != 0 or record.get("case") not in CASES:
                fail(f"invalid allocation record: {record}")
            output = P / record["output"]
            if sha(output) != record["output_sha256"]:
                fail(f"allocation output hash mismatch: {output}")
            parse_allocation(output)
    need(P / "allocation-summary.json")
    need(P / "allocation-comparisons.json")


def verify_profile() -> None:
    for phase in PHASES:
        out = P / "profile" / phase
        binding = json.loads(need(out / "binding.json").read_text())
        native = next(row for row in read(f"builds-{phase}.json") if row["label"] == "native")
        if binding.get("binary_sha256") != native["binary_sha256"]:
            fail(f"{phase} profile binary binding mismatch")
        commands = json.loads(need(out / "commands.json").read_text())
        if {row["name"] for row in commands} != {"record", "self-symbols", "inclusive-symbols", "counters-10", "counters-210", "rss"}:
            fail(f"{phase} profile command census mismatch")
        for row in commands:
            if row.get("exit_code") != 0:
                fail(f"{phase} profile command failed")
            if row["name"] not in {"self-symbols", "inclusive-symbols"}:
                if "prefix" not in row.get("command", []) or "commit" not in row.get("command", []):
                    fail(f"{phase} profile denominator is not the commit prefix")
            elif row["command"][:2] != ["perf", "report"]:
                fail(f"{phase} profile symbol report command mismatch")


def verify_tests_gates() -> None:
    inventory = read("functional-test-inventory.json")
    functional = read("functional-tests.json")
    tests = functional.get("tests", [])
    if len(tests) != 1:
        fail("candidate functional threshold test receipt is empty")
    row = tests[0]
    if row.get("exit_code") != 0 or row.get("source") != inventory["source"]:
        fail(f"invalid candidate functional test receipt: {row}")
    if row.get("filter") != inventory["filter"] or row.get("required_tests") != inventory["tests"]:
        fail("candidate functional test inventory binding mismatch")
    passed = row.get("passed_tests", [])
    minimum = len(inventory["tests"]) + inventory.get("minimum_additional_tests", 0)
    if row.get("passed_test_count") != len(passed) or len(passed) < minimum:
        fail("candidate functional test receipt omitted policy-intersection tests")
    source_path = ROOT / row["source"]
    if not source_path.exists() or sha(source_path) != row.get("source_sha256"):
        fail("candidate functional test source hash mismatch")
    log = P / row["log"]
    if sha(log) != row["log_sha256"]:
        fail(f"functional test log hash mismatch: {log}")
    log_text = log.read_text()
    for test_name in inventory["tests"]:
        if not any(name == test_name or name.endswith(f"::{test_name}") for name in passed):
            fail(f"functional test log lacks a passing line for {test_name}")
    if row.get("required_additional_tests") != inventory.get("required_additional_tests", []):
        fail("functional test receipt omitted the focused policy-test inventory")
    for test_name in inventory.get("required_additional_tests", []):
        if not any(name == test_name or name.endswith(f"::{test_name}") for name in passed):
            fail(f"functional test log lacks a passing policy test for {test_name}")
    actual = set(re.findall(r"^test\s+([^\s]+)\s+\.\.\.\s+ok$", log_text, flags=re.MULTILINE))
    if set(passed) != actual:
        fail("functional test receipt does not match the passing test lines")
    integration = read("integration/results.json")
    if len(integration) != 7 or {row["name"] for row in integration} != {"fmt", "check", "clippy", "tests-default", "tests", "facade", "rustdoc"}:
        fail("integration gate census mismatch")
    if any(row.get("exit_code") != 0 for row in integration):
        fail("an integration gate failed")
    evidence = read("evidence/results.json")
    if len(evidence) != 6 or any(row.get("exit_code") != 0 for row in evidence):
        fail("evidence gate census or status mismatch")
    current = source_map()
    for group, records in (("integration", integration), ("evidence", evidence)):
        for receipt in records:
            if any(receipt["source_sha256"].get(name) != digest for name, digest in current.items()):
                fail(f"{group} gate source binding mismatch: {receipt['name']}")
            log = need(P / group / (receipt["name"] + ".log"))
            if "log_sha256" in receipt and sha(log) != receipt["log_sha256"]:
                fail(f"{group} gate log hash mismatch")
    quality = read("quality-summary.json")
    if {row["name"] for row in quality} != {row["name"] for row in integration}:
        fail("quality gate census mismatch")
    for row in quality:
        log = P / "integration" / (row["name"] + ".log")
        if sha(log) != row["log_sha256"]:
            fail("quality log binding mismatch")
        matches = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;", log.read_text())
        totals = {name: sum(int(item[index]) for item in matches)
                  for index, name in enumerate(("passed", "failed", "ignored"))}
        if totals != row["test_totals"]:
            fail("quality test totals mismatch")


def verify_cleanup() -> None:
    cleanup = P / "cleanup.json"
    if cleanup.exists():
        data = read("cleanup.json")
        if {row["path"] for row in data["removed"]} != {str(path) for path in SCRATCH}:
            fail("cleanup scope is not exactly the five 0704-owned paths")
        if any(not row.get("removed") for row in data["removed"]):
            fail("cleanup left owned scratch behind")
        mechanism_receipts = [
            P / name
            for name in ("mechanism/build.json", "mechanism/mechanism-build.json", "mechanism/build-mechanism.json")
            if (P / name).is_file()
        ]
        expected_binaries = 7 + (1 if mechanism_receipts else 0)
        if len(data.get("binaries", [])) != expected_binaries:
            fail(f"cleanup did not retain all {expected_binaries} frozen binary hashes")
        witness = read("candidate-source-witness.json")
        handoff = historical_handoff(witness)
        retained = data.get("historical_worktree_witness", {})
        if retained.get("path") != handoff.get("candidate_worktree"):
            fail("cleanup did not retain the candidate worktree path witness")
        if retained.get("candidate_patch_sha256") != handoff.get("candidate_patch_sha256"):
            fail("cleanup did not retain the candidate patch witness hash")
        if retained.get("candidate_source_sha256") != handoff.get("candidate_source_sha256"):
            fail("cleanup did not retain the candidate source witness hashes")
        lock = ROOT / "Cargo.lock"
        if lock.exists() and data.get("workspace_cargo_lock_preserved_sha256") != sha(lock):
            fail("workspace Cargo.lock changed during cleanup")
    else:
        # Before cleanup, owned scratch may exist; no unrelated deletion check
        # is needed.  The terminal seal requires cleanup.json.
        for path in SCRATCH:
            if path.exists() and path.is_symlink():
                fail(f"owned scratch path is unexpectedly a symlink: {path}")


def main() -> None:
    verify_report_binding()
    verify_sources()
    verify_candidate_witness()
    verify_scripts_manifest()
    verify_probe_inputs()
    verify_builds()
    verify_observer()
    verify_mechanism()
    control = read("control-manifest.json")
    source = ROOT / control["source"]
    if sha(source) != control["source_sha256"]:
        fail("control source hash mismatch")
    control_path = P / control["control"]
    if control_path.exists() and sha(control_path) != control["control_sha256"]:
        fail("control archive hash mismatch")
    if len(control["members"]) != 103 or sum(row["replacements"] for row in control["members"]) != 43:
        fail("control member/replacement census mismatch")
    if CASES.keys() != EXPECTED_CASES:
        fail("native case census mismatch")
    verify_native()
    verify_refusal()
    verify_allocations()
    verify_profile()
    subprocess.run(["python3", str(P / "audit_followup.py")], cwd=ROOT, check=True)
    subprocess.run(["python3", str(P / "mechanism/audit.py")], cwd=ROOT, check=True)
    subprocess.run(["python3", str(P / "audit-statistics.py")], cwd=ROOT, check=True)
    verify_tests_gates()
    verify_cleanup()
    print("0704 audit PASS: PPTX-only source boundary, probes, observer, measurements, profiles, tests, gates, and cleanup")


if __name__ == "__main__":
    main()
