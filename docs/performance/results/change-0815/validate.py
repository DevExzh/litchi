"""Fail-closed aggregate validator for the 0815 workflow packet.

This reader is intentionally independent of the numerical and profile
readers.  It verifies their retained schemas and ``--check`` paths, joins
their custody and policy results, checks stage chronology, and optionally
requires the final six-binary cleanup witness.  It never invokes Cargo,
rustfmt, a probe, a workload, Valgrind, or a profiler, and it writes no
artifact of its own.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0815"
OLD_0813 = ROOT / "docs/performance/results/change-0813"
BASE = "55bb2ead3498043dd22b53555507b24402486248"
SOURCE_ALLOWLIST = {"crates/litchi-pptx/src/notes/codec.rs"}
SOURCE_COUNT = 9196
CASES = tuple(
    (shape, mode)
    for shape in ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
    for mode in ("capture", "commit", "lifecycle")
)
ORDERS = {
    "qualification": (("before",),),
    "native": (
        ("before", "after"),
        ("after", "before"),
        ("before", "after"),
        ("after", "before"),
        ("after", "before"),
        ("before", "after"),
    ),
    "allocation": (("before", "after"), ("after", "before")),
}
LANE_SAMPLES = {"qualification": 1, "native": 30, "allocation": 3}
LANE_REPORTS = {"qualification": 18, "native": 216, "allocation": 72}
LANE_TOTALS = {"qualification": 18, "native": 6480, "allocation": 216}
PROFILE_JOBS = ((0, "before"), (0, "after"), (1, "after"), (1, "before"))
READER_CHECKS = (
    ("analysis.py", "analysis.json"),
    ("root_audit.py", "root-audit.json"),
    ("profile_analysis.py", "profile-analysis.json"),
    ("codegen_analysis.py", "codegen-analysis.json"),
    ("quality_summary.py", "quality-summary.json"),
)


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    path = Path(value)
    if not path.is_absolute():
        path = P / path
    path = path.resolve()
    require(path.is_relative_to(P), f"{label}: path escapes packet")
    return path


def packet_artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    path = packet_path(value.get("path"), label)
    require(type(value.get("bytes")) is int and value["bytes"] >= 0,
            f"{label}: byte count is invalid")
    require(is_sha(value.get("sha256")), f"{label}: SHA-256 is invalid")
    require(path.is_file() and not path.is_symlink(), f"{label}: file is missing")
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: SHA-256 changed")
    return path


def sealed_external_artifact(value: Any, label: str, relative: str) -> Path:
    """Check an artifact outside this packet against the sealed 0813 index."""
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: artifact path is missing")
    path = Path(raw).resolve()
    seal_key = (relative if relative.startswith("docs/performance/results/change-0813/")
                else f"docs/performance/results/change-0813/{relative}")
    expected = (ROOT / seal_key).resolve()
    require(path == expected, f"{label}: sealed path changed")
    require(path.is_file() and not path.is_symlink(), f"{label}: file is missing")
    require(type(value.get("bytes")) is int and value["bytes"] >= 0,
            f"{label}: byte count is invalid")
    require(is_sha(value.get("sha256")), f"{label}: SHA-256 is invalid")
    require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
            f"{label}: artifact identity changed")
    seal = read(OLD_0813 / "seal.json")
    require(seal.get("files", {}).get(seal_key) == value["sha256"],
            f"{label}: artifact is not in the sealed 0813 index")
    return path


def external_artifact(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: binary path is missing")
    path = Path(raw).resolve()
    require(path.is_absolute() and path.parent == TARGET,
            f"{label}: binary path is outside owned target")
    size, digest = value.get("bytes"), value.get("sha256")
    require(type(size) is int and size > 0, f"{label}: binary byte count is invalid")
    require(is_sha(digest), f"{label}: binary SHA-256 is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size and sha(path) == digest,
                f"{label}: live binary identity changed")
        return {"path": str(path), "bytes": size, "sha256": digest}
    require(cleanup is not None and cleanup.get("target_removed") is True,
            f"{label}: missing binary has no cleanup witness")
    require(not TARGET.exists(), f"{label}: target remains after cleanup")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), f"{label}: cleanup binary list is malformed")
    matches = [item for item in removed if isinstance(item, dict)
               and item.get("path") == str(path)]
    require(len(matches) == 1 and matches[0].get("bytes") == size
            and matches[0].get("sha256") == digest,
            f"{label}: exact removed binary identity is missing")
    return {"path": str(path), "bytes": size, "sha256": digest}


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{label}: source manifest is malformed")
    revision, files = value.get("revision"), value.get("files")
    require(isinstance(revision, str) and len(revision) == 40,
            f"{label}: revision is malformed")
    require(isinstance(files, dict) and len(files) == SOURCE_COUNT,
            f"{label}: source census count changed")
    require(all(isinstance(name, str) and is_sha(digest)
                for name, digest in files.items()),
            f"{label}: source digest map is malformed")
    return {"revision": revision, "files": dict(files)}


def current_source() -> dict[str, str]:
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
    names = [name for name in raw.decode().split("\0") if name]
    return {name: sha(ROOT / name) for name in names}


def verify_origin() -> dict[str, Any]:
    origin = read(P / "origin.json")
    require(origin.get("schema") == "litchi.performance.0815.origin.v1",
            "origin schema changed")
    require(origin.get("base") == BASE and origin.get("worktree") == str(ROOT),
            "origin base or worktree changed")
    require(origin.get("historical_timing_pool") is False
            and origin.get("cross_format_lane") is False
            and origin.get("iwork") == "out of scope",
            "origin scope changed")
    production = origin.get("production_source")
    require(isinstance(production, dict)
            and production.get("revision") == BASE
            and production.get("tracked_file_count") == SOURCE_COUNT
            and production.get("candidate_allowlist") == sorted(SOURCE_ALLOWLIST),
            "origin production custody changed")
    unrelated = origin.get("unrelated")
    require(isinstance(unrelated, dict), "origin unrelated identities are missing")
    for name, digest in unrelated.items():
        path = ROOT / name
        require(path.is_file() and sha(path) == digest, f"unrelated file changed: {name}")
    architecture = read(P / "architecture-inputs.json")
    require(isinstance(architecture, dict), "architecture input map is malformed")
    for name, digest in architecture.items():
        path = ROOT / name
        require(path.is_file() and sha(path) == digest, f"architecture input changed: {name}")
    root_inputs = origin.get("root_inputs")
    require(isinstance(root_inputs, dict)
            and root_inputs.get("copies_captured_before_first_build") is True
            and is_sha(root_inputs.get("root-Cargo.lock"))
            and is_sha(root_inputs.get("rustfmt.toml")),
            "root input custody changed")
    for name, digest in {
        "Cargo.lock": root_inputs["root-Cargo.lock"],
        "rustfmt.toml": root_inputs["rustfmt.toml"],
    }.items():
        copy = P / "inputs" / ("root-Cargo.lock" if name == "Cargo.lock" else "rustfmt.toml")
        require(copy.is_file() and sha(copy) == digest, f"frozen {name} copy changed")
        require((ROOT / name).is_file() and sha(ROOT / name) == digest,
                f"live {name} changed")
    return {"origin": origin, "root_inputs": {
        "Cargo.lock": root_inputs["root-Cargo.lock"],
        "rustfmt.toml": root_inputs["rustfmt.toml"],
    }}


def verify_frozen_build_inputs() -> None:
    expected = {
        "plan.json", "adoption-policy.json", "analysis-plan.json", "custody.py",
        "build.py", "capture.py", "profile.py", "quality.py", "probe_quality.py",
        "quality_reuse.py", "apply_candidate.py", "restore_candidate.py", "origin.json", "host.json",
        "inheritance.json", "architecture-inputs.json", "inputs/root-Cargo.lock",
        "inputs/rustfmt.toml", "quality-reuse.json",
        "quality-reuse/reuse-inputs.json", "codegen.py", "protocol-review.md",
        "codegen_analysis.py", "source-review.md", "toolchain.json",
    }
    expected.update(
        str(path.relative_to(P))
        for path in (P / "candidate").rglob("*")
        if path.is_file() and not path.is_symlink()
    )
    architecture = read(P / "architecture-inputs.json")
    origin = read(P / "origin.json")
    unrelated = origin.get("unrelated")
    require(isinstance(architecture, dict) and isinstance(unrelated, dict),
            "frozen architecture or unrelated custody is missing")
    for leg in ("before", "after"):
        value = read(P / f"build-{leg}" / "frozen-inputs.json")
        require(value.get("schema") == "litchi.performance.0815.frozen-inputs.v1"
                and set(value) == {"schema", "packet", "root_inputs", "architecture", "unrelated"},
                f"{leg} frozen input envelope changed")
        require(set(value["packet"]) == expected, f"{leg} frozen input set changed")
        for name, digest in value["packet"].items():
            path = P / name
            require(is_sha(digest) and path.is_file() and sha(path) == digest,
                    f"{leg} frozen packet input changed: {name}")
        require(value["root_inputs"] == verify_origin()["root_inputs"],
                f"{leg} frozen root input receipt changed")
        require(value["architecture"] == architecture
                and value["unrelated"] == unrelated,
                f"{leg} frozen workspace custody changed")


def verify_plan() -> dict[str, Any]:
    plan = read(P / "plan.json")
    require(plan.get("schema") == "litchi.performance.0815.v1", "plan schema changed")
    require(plan.get("source_allowlist") == sorted(SOURCE_ALLOWLIST), "plan scope changed")
    require(plan.get("cases") == [{"mode": mode, "shape": shape} for shape, mode in CASES],
            "plan case order changed")
    for lane, orders in ORDERS.items():
        row = plan.get(lane)
        require(isinstance(row, dict) and row.get("orders") == [list(order) for order in orders],
                f"{lane} order changed")
        require(row.get("samples") == LANE_SAMPLES[lane]
                and row.get("reports") == LANE_REPORTS[lane]
                and row.get("samples_total") == LANE_TOTALS[lane],
                f"{lane} cardinality changed")
    require(plan["qualification"].get("binary") == "allocation",
            "qualification binary changed")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("owner") == "namespace_uri_probe::capture_region_0793"
            and profile.get("binary") == "profile"
            and profile.get("mode") == "capture"
            and profile.get("shape") == "large"
            and profile.get("orders") == [["before", "after"], ["after", "before"]]
            and profile.get("repeats") == 2
            and profile.get("samples") == 1
            and profile.get("warmup") == 0
            and profile.get("collect_at_start") is False
            and profile.get("events") == ["Ir"]
            and profile.get("expected_numbered_parts") == 1
            and profile.get("reports") == 4
            and profile.get("samples_total") == 4,
            "profile plan changed")
    require(plan.get("totals") == {"reports": 310, "samples": 6718},
            "plan totals changed")
    boot = plan.get("bootstrap")
    require(boot == {
        "resamples": 10000,
        "seed": 815815,
        "statistic": "median",
        "sorted_zero_based_endpoints": [250, 9749],
    }, "bootstrap plan changed")
    policy = read(P / "adoption-policy.json")
    require(policy.get("schema") == "litchi.performance.0815.adoption-policy.v1",
            "adoption policy schema changed")
    require(policy.get("frozen_before_build") is True
            and policy.get("useful_public_workflow_benefit_required") is True
            and policy.get("allocation_count_alone_sufficient") is False
            and policy.get("latency", {}).get("seed") == 815815
            and policy.get("latency", {}).get("resamples") == 10000
            and policy.get("benefit", {}).get("eligible_modes") == ["capture", "lifecycle"]
            and policy.get("benefit", {}).get("minimum_improvement_percent") == 3.0,
            "adoption policy changed")
    return plan


def candidate_sources(before: dict[str, Any]) -> dict[str, Any]:
    manifest = read(P / "candidate/manifest.json")
    require(manifest.get("schema") == "litchi.performance.0815.candidate-manifest.v1",
            "candidate manifest schema changed")
    require(manifest.get("base_commit") == BASE
            and manifest.get("production_path") in SOURCE_ALLOWLIST
            and manifest.get("production_source_changed") is False
            and manifest.get("production_adoption") is False,
            "candidate manifest scope changed")
    files = manifest.get("files")
    rows = list(files.values()) if isinstance(files, dict) else [manifest]
    require(len(rows) == 1 and {row.get("production_path") for row in rows} == SOURCE_ALLOWLIST,
            "candidate manifest file set changed")
    row = rows[0]
    before_path = packet_artifact(row["before"], "candidate before")
    after_path = packet_artifact(row["after"], "candidate after")
    require(before["files"][next(iter(SOURCE_ALLOWLIST))] == sha(before_path),
            "candidate before differs from build-before source")
    before_bytes, after_bytes = before_path.read_bytes(), after_path.read_bytes()
    scanner = b"fn scan_processed_xml("
    tests = b"#[cfg(test)]\nmod tests"
    before_scanner, after_scanner = before_bytes.find(scanner), after_bytes.find(scanner)
    before_tests, after_tests = before_bytes.find(tests), after_bytes.find(tests)
    require(min(before_scanner, after_scanner, before_tests, after_tests) >= 0,
            "candidate scanner/test boundaries are missing")
    require(before_bytes[:before_scanner] == after_bytes[:after_scanner]
            and before_bytes[before_tests:] == after_bytes[after_tests:],
            "candidate changed source outside scan_processed_xml")

    # The direct Result<Event> match is already retained production behavior;
    # this batch is limited to borrowing the two payloads in the existing
    # Start/Empty arms.  Check the six spellings explicitly so a broader
    # event-loop rewrite cannot hide inside the archive.
    before_scan = before_bytes[before_scanner:before_tests]
    after_scan = after_bytes[after_scanner:after_tests]
    replacements = (
        (b"Ok(Event::Start(element))", b"Ok(Event::Start(ref element))", 1),
        (b"Ok(Event::Empty(element))", b"Ok(Event::Empty(ref element))", 1),
        (b".push(&element)", b".push(element)", 2),
        (b"                    &element,", b"                    element,", 2),
    )
    transformed = before_scan
    for old, new, count in replacements:
        require(before_scan.count(old) == count and after_scan.count(old) == 0,
                f"candidate replacement source count changed: {old!r}")
        require(after_scan.count(new) == count,
                f"candidate replacement target source count changed: {new!r}")
        transformed = transformed.replace(old, new)
    require(transformed == after_scan,
            "candidate scanner differs outside the six borrowed-arm substitutions")

    # The differential/refusal oracles live in the unchanged test suffix.
    # Keep an explicit witness for both names so a future source relocation
    # cannot silently remove the independent comparison paths.
    for marker in (b"fn buffered_scan_oracle(", b"fn inspect_element_oracle("):
        require(before_bytes.count(marker) == after_bytes.count(marker) > 0,
                f"candidate oracle marker changed: {marker!r}")
    expected = dict(before["files"])
    expected[next(iter(SOURCE_ALLOWLIST))] = sha(after_path)
    packet_artifact(manifest["patch"], "candidate patch")
    return {"revision": before["revision"], "files": expected}


def verify_builds(cleanup: dict[str, Any] | None) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    builds: dict[str, Any] = {}
    for leg in ("before", "after"):
        directory = P / f"build-{leg}"
        manifest = read(directory / "build.json")
        require(manifest.get("schema") == f"litchi.performance.0815.build-{leg}.v1",
                f"{leg} build schema changed")
        source_path = packet_artifact(manifest["source"], f"{leg} source")
        require(source_path == (directory / "source.json").resolve(),
                f"{leg} source path changed")
        source = source_manifest(source_path, f"{leg} source")
        require(manifest.get("root_inputs") == verify_origin()["root_inputs"],
                f"{leg} root input receipt changed")
        require(manifest.get("architecture") == read(P / "architecture-inputs.json")
                and manifest.get("unrelated") == read(P / "origin.json").get("unrelated"),
                f"{leg} workspace custody changed")
        frozen_path = packet_artifact(manifest.get("frozen_inputs"), f"{leg} frozen inputs")
        require(frozen_path == (directory / "frozen-inputs.json").resolve(),
                f"{leg} frozen input path changed")
        rows = manifest.get("rows")
        require(isinstance(rows, list) and len(rows) == 3, f"{leg} build row count changed")
        last = -math.inf
        expected_features = {
            "native": [],
            "allocation": ["--features", "allocator-metrics"],
            "profile": ["--features", "capture-profile"],
        }
        seen_names = []
        for row in rows:
            require(row.get("name") in {"native", "allocation", "profile"}
                    and row.get("exit_code") == 0, f"{leg} build command failed")
            name = row["name"]
            require(name not in seen_names, f"{leg} build variant duplicated: {name}")
            seen_names.append(name)
            expected_command = [
                "cargo", "build", "--offline", "--locked", "--release",
                "--manifest-path", str(P / "probe-src/Cargo.toml"),
                *expected_features[name],
            ]
            require(row.get("command") == expected_command,
                    f"{leg} {name} build command changed")
            require(isinstance(row.get("started"), (int, float))
                    and isinstance(row.get("ended"), (int, float))
                    and last <= row["started"] <= row["ended"],
                    f"{leg} build chronology changed")
            last = row["ended"]
            packet_artifact(row["log"], f"{leg} build log")
        require(seen_names == ["native", "allocation", "profile"],
                f"{leg} build order changed")
        require(manifest.get("environment") == {
            "CARGO_TARGET_DIR": str(TARGET),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0",
        }, f"{leg} build environment changed")
        binaries = manifest.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
                f"{leg} binary matrix changed")
        checked = {name: external_artifact(value, f"{leg} {name}", cleanup)
                   for name, value in binaries.items()}
        builds[leg] = {"manifest": manifest, "source": source, "binaries": checked}
    before, after = builds["before"]["source"], builds["after"]["source"]
    require(before["revision"] == BASE, "before source revision changed")
    changed = {name for name in before["files"] | after["files"]
               if before["files"].get(name) != after["files"].get(name)}
    require(changed == SOURCE_ALLOWLIST, f"build source scope changed: {changed}")
    return builds, before, after


def verify_codegen(builds: dict[str, Any]) -> None:
    """Check the retained ordinary/profile scanner disassembly witnesses."""
    for leg in ("before", "after"):
        value = read(P / f"codegen-{leg}/receipt.json")
        require(value.get("schema") == "litchi.performance.0815.codegen.v1"
                and value.get("leg") == leg,
                f"{leg} codegen receipt changed")
        rows = value.get("rows")
        require(isinstance(rows, list) and len(rows) == 2
                and [row.get("variant") for row in rows] == ["native", "profile"],
                f"{leg} codegen variant set changed")
        for row in rows:
            variant = row["variant"]
            binary = builds[leg]["manifest"]["binaries"][variant]
            require(row.get("binary") == binary,
                    f"{leg} {variant} codegen binary changed")
            require(row.get("source") == builds[leg]["manifest"]["source"],
                    f"{leg} {variant} codegen source changed")
            address, size = row.get("address"), row.get("size")
            require(type(address) is int and address >= 0 and type(size) is int and size > 0,
                    f"{leg} {variant} scanner symbol range changed")
            require(isinstance(row.get("symbol"), str)
                    and "scan_processed_xml" in row["symbol"],
                    f"{leg} {variant} scanner symbol changed")
            nm = row.get("nm")
            require(isinstance(nm, dict) and nm.get("exit_code") == 0
                    and nm.get("matched_count") == 1,
                    f"{leg} {variant} nm receipt changed")
            packet_artifact(nm.get("symbols"), f"{leg} {variant} nm symbols")
            objdump = row.get("objdump")
            require(isinstance(objdump, dict) and objdump.get("exit_code") == 0,
                    f"{leg} {variant} objdump receipt changed")
            assembly = packet_artifact(objdump.get("assembly"),
                                       f"{leg} {variant} scanner assembly")
            require(assembly.read_text(encoding="utf-8", errors="replace").strip(),
                    f"{leg} {variant} scanner assembly is empty")
        tools = value.get("tools")
        require(isinstance(tools, dict) and set(tools) == {"nm", "objdump"}
                and all(isinstance(row, dict) and row.get("exit_code") == 0
                        for row in tools.values()),
                f"{leg} codegen tool versions changed")


def verify_codegen_analysis(builds: dict[str, Any]) -> dict[str, Any]:
    """Verify the exact arm-copy replay against the retained assembly.

    The four rows are deliberately checked independently of the parser that
    produced ``codegen-analysis.json``.  The validator recomputes the bounded
    vector-copy pattern and the reader-to-dispatch window from each retained
    assembly artifact, then binds the JSON result to those recomputed values.
    """
    value = read(P / "codegen-analysis.json")
    require(value.get("schema") == "litchi.performance.0815.codegen-analysis.v1",
            "codegen analysis schema changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 4,
            "codegen analysis row count changed")
    expected = [(leg, variant) for leg in ("before", "after")
                for variant in ("native", "profile")]
    require([(row.get("leg"), row.get("variant")) for row in rows] == expected,
            "codegen analysis row order changed")
    def instructions_from(assembly: Path, address: int, size: int) -> list[dict[str, Any]]:
        result = []
        for line in assembly.read_text(encoding="utf-8", errors="replace").splitlines():
            match = re.fullmatch(r"\s*([0-9a-f]+):\s+(.+)", line)
            if match is None:
                continue
            absolute = int(match.group(1), 16)
            require(address <= absolute < address + size,
                    f"assembly instruction escapes scanner symbol: {assembly}")
            result.append({"offset": absolute - address, "instruction": match.group(2)})
        require(result, f"assembly contains no instructions: {assembly}")
        return result

    def arm_copies(instructions: list[dict[str, Any]]) -> list[dict[str, Any]]:
        result = []
        for index in range(len(instructions) - 3):
            group = instructions[index:index + 4]
            first = re.fullmatch(
                r"movups\s+0x8\((%[a-z0-9]+)\),(%xmm[0-9]+)",
                group[0]["instruction"],
            )
            second = re.fullmatch(
                r"movups\s+0x18\((%[a-z0-9]+)\),(%xmm[0-9]+)",
                group[1]["instruction"],
            )
            if first is None or second is None or first.group(1) != second.group(1):
                continue
            stores = [
                re.fullmatch(r"mov(?:aps|ups)\s+(%xmm[0-9]+),(.+)",
                             item["instruction"])
                for item in group[2:]
            ]
            if not all(stores) or {item.group(1) for item in stores} != {
                first.group(2), second.group(2)
            }:
                continue
            result.append({"offset": group[0]["offset"], "bytes": 32,
                           "instructions": group})
        return result

    def reader_window(instructions: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], bool]:
        calls = [
            index for index, item in enumerate(instructions)
            if "call" in item["instruction"]
            and "<quick_xml::reader::Reader<R>::read_event_impl>" in item["instruction"]
        ]
        require(len(calls) == 1, "scanner assembly reader call count changed")
        window = []
        terminated = False
        for item in instructions[calls[0]:calls[0] + 81]:
            window.append(item)
            if re.match(r"jmp\s+\*", item["instruction"]):
                terminated = True
                break
        return window, terminated

    for row in rows:
        key = (row["leg"], row["variant"])
        require(row.get("binary") == builds[row["leg"]]["manifest"]["binaries"][row["variant"]],
                f"codegen analysis binary changed: {key}")
        receipt = packet_artifact(row.get("receipt"), f"codegen analysis receipt {key}")
        require(receipt == (P / f"codegen-{row['leg']}/receipt.json").resolve(),
                f"codegen analysis receipt path changed: {key}")
        codegen = read(receipt)
        codegen_row = next(
            (item for item in codegen["rows"] if item.get("variant") == row["variant"]),
            None,
        )
        require(isinstance(codegen_row, dict), f"codegen row is missing: {key}")
        assembly = packet_artifact(codegen_row["objdump"]["assembly"],
                                   f"codegen analysis assembly {key}")
        instructions = instructions_from(assembly, codegen_row["address"],
                                         codegen_row["size"])
        expected_arm = arm_copies(instructions)
        expected_window, terminated = reader_window(instructions)
        expected_moves = [
            item for item in instructions
            if re.match(r"(?:v?movups|v?movaps|v?movdqu|v?movdqa)\s",
                        item["instruction"])
        ]
        require(row.get("symbol_bytes") == codegen_row["size"],
                f"codegen analysis symbol size changed: {key}")
        require(row.get("instructions") == len(instructions),
                f"codegen analysis instruction count changed: {key}")
        require(row.get("arm_copy_sequences") == expected_arm,
                f"codegen analysis arm-copy replay changed: {key}")
        require(row.get("all_vector_moves") == expected_moves,
                f"codegen analysis vector-move replay changed: {key}")
        require(row.get("reader_to_indirect_dispatch") == expected_window
                and row.get("indirect_dispatch_found") is True and terminated is True,
                f"codegen analysis reader window changed: {key}")
        require(row.get("vector_moves_in_window") == 0,
                f"codegen analysis pre-dispatch vector count changed: {key}")
        require(len(expected_arm) == 2 if row["leg"] == "before" else len(expected_arm) < 2,
                f"codegen analysis arm-copy count changed: {key}")
    return value


def verify_codegen_gate(analysis: dict[str, Any], native_start: float) -> None:
    """Require the explicit root-reviewed mechanism gate before fresh capture."""
    value = read(P / "codegen-gate.json")
    require(value.get("schema") == "litchi.performance.0815.codegen-gate.v1"
            and value.get("advance") is True
            and value.get("reviewed_after_code_generation") is True,
            "codegen gate does not authorize the planned capture")
    reason = next(
        (value.get(name) for name in ("reason", "review_reason", "mechanism_reason")
         if isinstance(value.get(name), str) and value.get(name).strip()),
        None,
    )
    require(reason is not None, "codegen gate root review reason is missing")
    analysis_ref = value.get("analysis")
    require(packet_artifact(analysis_ref, "codegen gate analysis")
            == (P / "codegen-analysis.json").resolve(),
            "codegen gate analysis binding changed")
    receipts = value.get("receipts")
    require(isinstance(receipts, dict) and set(receipts) == {"before", "after"},
            "codegen gate receipt binding changed")
    expected = {
        "before": (P / "codegen-before/receipt.json").resolve(),
        "after": (P / "codegen-after/receipt.json").resolve(),
    }
    for leg, expected_path in expected.items():
        top_level = value.get(leg)
        require(isinstance(top_level, dict) and top_level == receipts[leg],
                f"codegen gate {leg} descriptor disagrees with receipts")
        require(packet_artifact(top_level, f"codegen gate {leg} receipt") == expected_path,
                f"codegen gate {leg} receipt binding changed")
    analysis_rows = {(row["leg"], row["variant"]): row for row in analysis["rows"]}
    require(len(analysis_rows) == 4, "codegen gate analysis does not bind four rows")
    for row in analysis["rows"]:
        require(packet_artifact(row["receipt"], "codegen gate analysis row receipt")
                == (P / f"codegen-{row['leg']}/receipt.json").resolve(),
                "codegen gate row receipt binding changed")
    ended = value.get("ended")
    require(isinstance(ended, (int, float)) and math.isfinite(ended)
            and ended <= native_start,
            "codegen gate ended after native capture started")
    counts = value.get("arm_copy_counts")
    if counts is not None:
        require(counts == {
            "before": {variant: 2 for variant in ("native", "profile")},
            "after": {
                variant: next(
                    len(row["arm_copy_sequences"])
                    for row in analysis["rows"]
                    if row["leg"] == "after" and row["variant"] == variant
                )
                for variant in ("native", "profile")
            },
        }, "codegen gate arm-copy count binding changed")


def verify_application(before: dict[str, Any], after: dict[str, Any]) -> None:
    application = read(P / "application.json")
    require(application.get("schema") == "litchi.performance.0815.application.v1",
            "application schema changed")
    require(application.get("source_before", {}).get("files") == before["files"]
            and application.get("source", {}).get("files") == after["files"]
            and application.get("allowlist") == sorted(SOURCE_ALLOWLIST),
            "application source custody changed")
    packet_artifact(application["manifest"], "application manifest")
    packet_artifact(application["patch"], "application patch")


def verify_qualification_audit(before: dict[str, Any]) -> None:
    """Require the independent before-only replay before candidate use."""
    value = read(P / "qualification-audit.json")
    require(value.get("schema") == "litchi.performance.0815.qualification-audit.v1"
            and value.get("passed") is True
            and value.get("accepted_before_application") is True
            and value.get("application_absent_at_acceptance") is True
            and value.get("build_after_absent_at_acceptance") is True,
            "qualification audit acceptance witness changed")
    require(value.get("reports") == 18 and value.get("samples") == 18
            and value.get("timings_imported") is False
            and value.get("all_semantic_oracles_checked") is True,
            "qualification audit cardinality or oracle witness changed")
    source = value.get("before_source")
    require(isinstance(source, dict)
            and source.get("revision") == before["revision"]
            and source.get("files") == before["files"],
            "qualification audit source differs from baseline")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 18,
            "qualification audit rows changed")
    require({(row.get("shape"), row.get("mode")) for row in rows}
            == set(CASES), "qualification audit case set changed")
    historical = value.get("historical_oracle")
    require(isinstance(historical, dict)
            and historical.get("timings_imported") is False
            and len(historical.get("rows", [])) == 18,
            "qualification audit historical-oracle witness changed")


def verify_lane(lane: str, plan: dict[str, Any], builds: dict[str, Any],
                expected_source: dict[str, Any], cleanup: dict[str, Any] | None) -> tuple[float, float]:
    directory = P / lane
    complete = read(directory / "complete.json")
    require(complete.get("schema") == f"litchi.performance.0815.{lane}.complete.v1",
            f"{lane} completion schema changed")
    require(complete.get("children") == LANE_REPORTS[lane]
            and complete.get("reports") == LANE_REPORTS[lane]
            and complete.get("samples") == LANE_TOTALS[lane]
            and complete.get("plan_sha256") == sha(P / "plan.json"),
            f"{lane} completion cardinality changed")
    source_path = packet_artifact(complete["source"], f"{lane} source")
    require(source_manifest(source_path, f"{lane} source") == expected_source,
            f"{lane} source differs from expected build")
    receipts = read(packet_artifact(complete["receipts"], f"{lane} receipts"))
    expected_jobs = [
        (block, shape, mode, leg)
        for block, order in enumerate(ORDERS[lane])
        for shape, mode in CASES
        for leg in order
    ]
    require(len(receipts) == len(expected_jobs), f"{lane} receipt count changed")
    first, last = math.inf, -math.inf
    previous = -math.inf
    for index, (row, identity) in enumerate(zip(receipts, expected_jobs)):
        block, shape, mode, leg = identity
        require(row.get("schema") == "litchi.performance.0815.capture-receipt.v1"
                and (row.get("lane"), row.get("block"), row.get("shape"),
                     row.get("mode"), row.get("leg"))
                == (lane, block, shape, mode, leg)
                and row.get("exit_code") == 0,
                f"{lane} receipt identity changed: {index}")
        require(row.get("root_inputs") == verify_origin()["root_inputs"],
                f"{lane} root input receipt changed: {index}")
        require(row.get("architecture") == read(P / "architecture-inputs.json")
                and row.get("unrelated") == read(P / "origin.json").get("unrelated"),
                f"{lane} workspace custody changed: {index}")
        started, ended = row.get("started"), row.get("ended")
        require(isinstance(started, (int, float)) and isinstance(ended, (int, float))
                and previous <= started <= ended,
                f"{lane} receipt chronology changed: {index}")
        first, last = min(first, started), max(last, ended)
        previous = ended
        kind = "allocation" if lane == "qualification" else lane
        expected_binary = builds[leg]["binaries"][kind]
        require(row.get("binary") == expected_binary, f"{lane} binary changed: {index}")
        packet_artifact(row["report"], f"{lane} report {index}")
        packet_artifact(row["log"], f"{lane} log {index}")
        packet_artifact(row["rss"], f"{lane} RSS {index}")
    return first, last


def verify_quality_summary(expected_after: dict[str, Any], root_inputs: dict[str, str]) -> None:
    value = read(P / "quality-summary.json")
    require(value.get("schema") == "litchi.performance.0815.quality-summary.v1",
            "quality summary schema changed")
    production = value.get("production_after")
    architecture = read(P / "architecture-inputs.json")
    unrelated = read(P / "origin.json").get("unrelated")
    require(isinstance(production, dict) and production.get("root_inputs") == root_inputs,
            "quality summary root-input custody changed")
    quality_receipt = read(P / "quality-after.json")
    require(quality_receipt.get("schema") == "litchi.performance.0815.quality-after.v1"
            and quality_receipt.get("root_inputs") == root_inputs
            and quality_receipt.get("architecture") == architecture
            and quality_receipt.get("unrelated") == unrelated,
            "production quality receipt custody changed")
    require(production.get("schema") == quality_receipt.get("schema"),
            "production quality schema binding changed")
    require(production.get("source") == quality_receipt.get("source")
            and production.get("checks") == quality_receipt.get("checks"),
            "production quality summary receipt bindings changed")
    source_path = packet_artifact(production["source"], "quality summary source")
    require(source_path == packet_artifact(quality_receipt["source"],
                                           "quality receipt source"),
            "quality summary source artifact binding changed")
    require(source_manifest(source_path, "quality summary source") == expected_after,
            "quality summary source differs from candidate")
    checks_path = packet_artifact(production["checks"], "quality summary checks")
    require(checks_path == packet_artifact(quality_receipt["checks"],
                                           "quality receipt checks"),
            "quality summary checks artifact binding changed")
    gates = production.get("gates")
    require(isinstance(gates, list) and len(gates) == 6
            and all(row.get("status") == "pass" and row.get("exit_code") == 0 for row in gates),
            "quality summary production gates changed")
    tests = production.get("tests")
    require(isinstance(tests, dict) and tests.get("suites") == 85
            and tests.get("passed") == 1241 and tests.get("failed") == 0
            and tests.get("ignored") == 3, "production test summary changed")
    for leg in ("before", "after"):
        probe = value.get("probe", {}).get(leg)
        require(isinstance(probe, dict), f"probe {leg} quality summary changed")
        inputs_path = packet_artifact(probe.get("inputs"), f"probe {leg} inputs")
        require(inputs_path == (P / f"probe-quality-{leg}/inputs.json").resolve(),
                f"probe {leg} inputs path changed")
        inputs = read(inputs_path)
        require(inputs.get("root_inputs") == root_inputs
                and inputs.get("architecture") == architecture
                and inputs.get("unrelated") == unrelated,
                f"probe {leg} input custody changed")
        receipts_path = packet_artifact(probe.get("receipts"), f"probe {leg} receipts")
        require(receipts_path == (P / f"probe-quality-{leg}/receipts.json").resolve(),
                f"probe {leg} receipts path changed")
        receipts = read(receipts_path)
        require(isinstance(receipts, list) and len(receipts) == 3,
                f"probe {leg} receipt count changed")
        gates = probe.get("gates")
        require(isinstance(gates, list) and len(gates) == 3
                and all(row.get("status") == "pass" and row.get("exit_code") == 0
                        for row in gates),
                f"probe {leg} quality gates changed")
        require(probe.get("tests", {}).get("suites") == 1
                and probe.get("tests", {}).get("passed") == 36
                and probe.get("tests", {}).get("failed") == 0
                and probe.get("tests", {}).get("ignored") == 0,
                f"probe {leg} test summary changed")


def verify_quality_reuse(expected_before: dict[str, Any], root_inputs: dict[str, str]) -> None:
    """Verify that the exact-source baseline reused sealed 0813 gates."""
    value = read(P / "quality-reuse.json")
    require(value.get("schema") == "litchi.performance.0815.quality-reuse.v1",
            "quality reuse schema changed")
    require(value.get("mode") == "reuse-sealed-0813-after-production-gates"
            and value.get("source_files_equal") is True
            and value.get("cargo_executed") is False
            and value.get("gate_count") == 6
            and value.get("root_inputs") == root_inputs,
            "quality reuse contract changed")
    require(value.get("architecture") == read(P / "architecture-inputs.json")
            and value.get("unrelated") == read(P / "origin.json").get("unrelated"),
            "quality reuse workspace custody changed")
    source_path = sealed_external_artifact(
        value.get("source"), "quality reuse source",
        "docs/performance/results/change-0813/build-after/source.json",
    )
    source = source_manifest(source_path, "quality reuse source")
    require(source["files"] == expected_before["files"],
            "sealed 0813 after source differs from 0815 before source")
    reference_path = sealed_external_artifact(
        value.get("reference"), "quality reuse summary",
        "docs/performance/results/change-0813/quality-summary.json",
    )
    prior = read(reference_path)
    require(prior.get("schema") == "litchi.performance.0813.quality-summary.v1",
            "sealed 0813 quality summary schema changed")
    production = prior.get("production_after")
    require(isinstance(production, dict), "sealed 0813 production summary is missing")
    require(production.get("schema") == "litchi.performance.0813.quality-after.v1",
            "sealed 0813 production quality schema changed")
    gates = value.get("gates")
    require(gates == production.get("gates") and isinstance(gates, list)
            and len(gates) == 6
            and all(row.get("status") == "pass" and row.get("exit_code") == 0
                    for row in gates),
            "quality reuse gate identities changed")
    for index, gate in enumerate(gates, 1):
        raw = gate.get("log", {}).get("path") if isinstance(gate.get("log"), dict) else None
        require(isinstance(raw, str), f"quality reuse gate {index} log is missing")
        sealed_external_artifact(gate["log"], f"quality reuse gate {index} log",
                                 str(Path(raw).resolve().relative_to(ROOT)))
    tests = value.get("tests")
    require(isinstance(tests, dict)
            and tests.get("suites") == 85
            and tests.get("passed") == 1241
            and tests.get("failed") == 0
            and tests.get("ignored") == 3,
            "quality reuse test summary changed")
    inputs = packet_artifact(value.get("inputs"), "quality reuse inputs")
    reuse = read(inputs)
    require(reuse.get("current_source", {}).get("files") == expected_before["files"]
            and reuse.get("sealed_0813_source_files_equal") is True
            and reuse.get("root_inputs") == root_inputs,
            "quality reuse input custody changed")
    source_ref = reuse.get("sealed_0813_source")
    sealed_external_artifact(source_ref, "reuse input source",
                             "docs/performance/results/change-0813/build-after/source.json")
    summary_ref = reuse.get("sealed_0813_summary")
    sealed_external_artifact(summary_ref, "reuse input summary",
                             "docs/performance/results/change-0813/quality-summary.json")
    require(reuse.get("architecture") == read(P / "architecture-inputs.json"),
            "quality reuse architecture custody changed")
    for name, digest in reuse.get("architecture", {}).items():
        require(sha(ROOT / name) == digest, f"quality reuse architecture changed: {name}")
    unrelated = reuse.get("unrelated")
    require(isinstance(unrelated, dict), "quality reuse unrelated custody missing")
    for name, digest in unrelated.items():
        require(sha(ROOT / name) == digest, f"quality reuse unrelated file changed: {name}")
    probe = reuse.get("probe")
    expected_probe = {
        str(path.relative_to(P / "probe-src")): sha(path)
        for path in (P / "probe-src").rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    require(probe == expected_probe, "quality reuse probe custody changed")
    return None


def case_keys(rows: Iterable[dict[str, Any]], *, field: str = "case") -> set[str]:
    result = set()
    for row in rows:
        if field in row:
            result.add(row[field])
        else:
            result.add(f"{row.get('shape')}/{row.get('mode')}")
    return result


def median(values: list[int | float]) -> int | float:
    """Return the even-count median used by the independent raw audit."""
    require(values, "cannot take the median of an empty list")
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def verify_native_numerical_agreement(
    analysis: dict[str, Any], audit: dict[str, Any]
) -> None:
    """Require every retained native p50 pair to match the independent audit."""
    main_rows = analysis.get("native", {}).get("analysis", {}).get(
        "paired_by_block_before_after"
    )
    audit_rows = audit.get("native")
    require(isinstance(main_rows, dict), "workflow native paired rows are missing")
    require(isinstance(audit_rows, list), "root native paired rows are missing")
    main_cases = set(main_rows)
    audit_cases = {f"{row.get('shape')}/{row.get('mode')}" for row in audit_rows}
    require(main_cases == audit_cases,
            "main/root native numerical case set disagreement")
    require(len(main_cases) == len(CASES),
            "native numerical case count changed")
    audit_by_case = {
        f"{row.get('shape')}/{row.get('mode')}": row for row in audit_rows
    }
    require(len(audit_by_case) == len(audit_rows),
            "root native numerical cases are duplicated")
    for case in sorted(main_cases):
        main_case = main_rows[case]
        root_case = audit_by_case[case]
        metrics = main_case.get("metrics", {}).get("p50")
        require(isinstance(metrics, dict), f"{case}: native p50 metrics are missing")
        blocks = metrics.get("by_block")
        require(isinstance(blocks, list) and len(blocks) == 6,
                f"{case}: native block rows changed")
        root_blocks = root_case.get("blocks")
        require(isinstance(root_blocks, list) and len(root_blocks) == 6,
                f"{case}: root native block rows changed")
        root_by_block = {row.get("block"): row for row in root_blocks}
        require(set(root_by_block) == {0, 1, 2, 3, 4, 5},
                f"{case}: root native block identities changed")
        main_by_block = {row.get("block"): row for row in blocks}
        require(set(main_by_block) == set(root_by_block),
                f"{case}: main/root native block identities disagree")
        main_ratios = []
        for block in range(6):
            left, right = main_by_block[block], root_by_block[block]
            for field in ("before", "after", "ratio"):
                require(left.get(field) == right.get(field),
                        f"{case} block {block}: native {field} disagreement")
            main_ratios.append(left["ratio"])
        require(root_case.get("paired_ratios") == main_ratios,
                f"{case}: native paired ratio list disagreement")
        require(metrics.get("ratio_median") == root_case.get("ratio"),
                f"{case}: native ratio median disagreement")
        bootstrap = metrics.get("bootstrap")
        ci = root_case.get("ci95")
        require(isinstance(bootstrap, dict) and isinstance(ci, dict),
                f"{case}: native bootstrap CI is missing")
        require(bootstrap.get("ci_low") == ci.get("low")
                and bootstrap.get("ci_high") == ci.get("high")
                and bootstrap.get("resamples") == ci.get("resamples")
                and bootstrap.get("seed") == ci.get("seed"),
                f"{case}: native bootstrap CI disagreement")
        require(root_case.get("before_p50_ns") == median(
            [row["before"] for row in blocks]
        ) and root_case.get("after_p50_ns") == median(
            [row["after"] for row in blocks]
        ), f"{case}: native aggregate p50 disagreement")


ALLOCATION_METRICS = (
    "allocation_calls",
    "allocated_bytes",
    "net_live",
    "peak_above_entry",
)


def verify_allocation_numerical_agreement(
    analysis: dict[str, Any], audit: dict[str, Any]
) -> None:
    """Require all four per-block allocation guard metrics to match exactly."""
    main_rows = analysis.get("allocation", {}).get("analysis", {}).get(
        "paired_by_block_before_after"
    )
    audit_rows = audit.get("allocation")
    require(isinstance(main_rows, dict), "workflow allocation paired rows are missing")
    require(isinstance(audit_rows, list), "root allocation paired rows are missing")
    require(set(main_rows) == {f"{shape}/{mode}" for shape, mode in CASES},
            "workflow allocation numerical case set changed")
    expected = {
        (case, block, metric)
        for case in main_rows
        for block in (0, 1)
        for metric in ALLOCATION_METRICS
    }
    root_by_key = {
        (f"{row.get('shape')}/{row.get('mode')}", row.get("block"), row.get("metric")): row
        for row in audit_rows
    }
    require(len(root_by_key) == len(audit_rows),
            "root allocation numerical rows are duplicated")
    require(set(root_by_key) == expected,
            "main/root allocation numerical key set disagreement")
    for case, block, metric in sorted(expected):
        main_metric = main_rows[case].get("metrics", {}).get(metric)
        require(isinstance(main_metric, dict),
                f"{case} block {block}: allocation {metric} is missing")
        raw_blocks = main_metric.get("by_block", [])
        main_blocks = {row.get("block"): row for row in raw_blocks}
        require(len(main_blocks) == len(raw_blocks),
                f"{case}: allocation {metric} block rows are duplicated")
        require(set(main_blocks) == {0, 1},
                f"{case}: allocation {metric} block rows changed")
        left = main_blocks[block]
        right = root_by_key[(case, block, metric)]
        for field in ("before", "after"):
            require(left.get(field) == right.get(field),
                    f"{case} block {block} {metric}: {field} disagreement")
        require(right.get("increase") is (right.get("after") > right.get("before")),
                f"{case} block {block} {metric}: increase witness changed")


def verify_reader_agreement() -> dict[str, Any]:
    analysis = read(P / "analysis.json")
    audit = read(P / "root-audit.json")
    profile = read(P / "profile-analysis.json")
    require(analysis.get("schema") == "litchi.performance.0815.workflow-analysis.v1",
            "workflow analysis schema changed")
    require(audit.get("schema") == "litchi.performance.0815.root-audit.v1",
            "root audit schema changed")
    require(profile.get("schema") == "litchi-0815-callgrind-profile-analysis-v1",
            "profile analysis schema changed")
    require(analysis.get("counts") == {
        "reports": 306,
        "samples": 6714,
        "native_reports": 216,
        "allocation_reports": 72,
        "qualification_reports": 18,
    }, "workflow analysis counts changed")
    require(audit.get("reports") == 306 and audit.get("samples") == 6714,
            "root audit counts changed")
    require(profile.get("summary", {}).get("profile_count") == 4
            and profile.get("summary", {}).get("all_four_profiles_complete") is True,
            "profile analysis count changed")
    decision = analysis.get("decision_guards")
    require(isinstance(decision, dict), "workflow decision guards are missing")
    audit_latency = case_keys(audit.get("latency_violations", []))
    main_latency = case_keys(decision.get("latency_violations", []))
    audit_benefits = case_keys(audit.get("benefits", []))
    main_benefits = case_keys(decision.get("eligible_benefits", []))
    audit_resources = {
        (f"{row.get('shape')}/{row.get('mode')}", row.get("block"), row.get("metric"))
        for row in audit.get("resource_violations", [])
    }
    main_resources = {
        (row.get("case"), row.get("block"), row.get("metric"))
        for row in decision.get("resource_violations", [])
    }
    require(audit_latency == main_latency, "main/root latency policy disagreement")
    require(audit_benefits == main_benefits, "main/root benefit policy disagreement")
    require(audit_resources == main_resources, "main/root resource policy disagreement")
    require(audit.get("adoption_eligible") == decision.get("adoption_eligible"),
            "main/root adoption eligibility disagreement")
    verify_native_numerical_agreement(analysis, audit)
    verify_allocation_numerical_agreement(analysis, audit)
    require("production_adoption" not in audit,
            "raw audit must not claim production retention")
    require(analysis.get("source_disposition_contract", {}).get("retention_decision_external") is True,
            "analysis disposition boundary changed")
    return {"analysis": analysis, "audit": audit, "profile": profile}


def verify_decision_and_disposition(before: dict[str, Any], after: dict[str, Any],
                                   numeric: dict[str, Any]) -> dict[str, Any]:
    decision = read(P / "decision.json")
    require(decision.get("schema") == "litchi.performance.0815.decision.v1",
            "decision schema changed")
    eligible = numeric["analysis"]["decision_guards"]["adoption_eligible"]
    require(decision.get("adoption_eligible") == eligible
            and isinstance(decision.get("production_adoption"), bool),
            "decision policy result changed")
    disposition = read(P / "disposition.json")
    require(disposition.get("schema") == "litchi.performance.0815.disposition.v1",
            "disposition schema changed")
    retained = disposition.get("production_change_retained")
    require(isinstance(retained, bool)
            and disposition.get("status") in {"retained", "rejected"}
            and retained is (disposition["status"] == "retained")
            and decision["production_adoption"] is retained,
            "decision/disposition retention mismatch")
    if retained:
        require(decision.get("adoption_eligible") is True,
                "retained disposition is numerically ineligible")
    expected = after if retained else before
    require(current_source() == expected["files"],
            "live source does not match final disposition")
    if retained:
        require("restored_source" not in disposition,
                "retained disposition has a restoration witness")
    else:
        restored = packet_artifact(disposition.get("restored_source"), "restored source")
        require(source_manifest(restored, "restored source") == before,
                "restored source differs from baseline")
    return disposition


def verify_cleanup(final: bool, builds: dict[str, Any], disposition: dict[str, Any]) -> dict[str, Any] | None:
    path = P / "cleanup.json"
    if not path.is_file():
        require(not final, "final validation requires cleanup.json")
        require(TARGET.is_dir() and not TARGET.is_symlink(),
                "owned target disappeared before cleanup witness")
        return None
    cleanup = read(path)
    require(cleanup.get("schema") == "litchi.performance.0815.cleanup.v1"
            and cleanup.get("target") == str(TARGET)
            and cleanup.get("target_removed") is True
            and not TARGET.exists(), "cleanup witness is incomplete")
    require(type(cleanup.get("removed_files")) is int and cleanup["removed_files"] >= 6
            and type(cleanup.get("removed_logical_bytes")) is int
            and cleanup["removed_logical_bytes"] >= 0
            and cleanup.get("binaries_verified_before_removal") is True,
            "cleanup counts changed")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 6,
            "cleanup binary witness count changed")
    expected = {
        (item["path"], item["bytes"], item["sha256"])
        for build in builds.values() for item in build["manifest"]["binaries"].values()
    }
    actual = {(item.get("path"), item.get("bytes"), item.get("sha256"))
              for item in removed if isinstance(item, dict)}
    require(actual == expected, "cleanup binary identities changed")
    for item in removed:
        require(not Path(item["path"]).exists(),
                f"removed binary remains: {item['path']}")
    source_leg = "after" if disposition["production_change_retained"] else "before"
    require(cleanup.get("source_leg") == source_leg,
            "cleanup source leg differs from disposition")
    source_path = packet_artifact(cleanup.get("source_manifest"), "cleanup source manifest")
    expected_source = builds[source_leg]["manifest"]["source"]
    require(source_path == packet_path(expected_source["path"], "cleanup expected source"),
            "cleanup source manifest path changed")
    return cleanup


def interval(rows: list[dict[str, Any]], label: str) -> tuple[float, float]:
    require(rows, f"{label}: no timestamped rows")
    first, last = math.inf, -math.inf
    previous = -math.inf
    for index, row in enumerate(rows):
        started, ended = row.get("started"), row.get("ended")
        require(isinstance(started, (int, float)) and math.isfinite(started)
                and isinstance(ended, (int, float)) and math.isfinite(ended)
                and previous <= started <= ended,
                f"{label}: chronology changed at {index}")
        first, last, previous = min(first, started), max(last, ended), ended
    return first, last


def timestamp_rows(path: Path, key: str) -> list[dict[str, Any]]:
    value = read(path)
    rows = value if key == "self" else value.get(key)
    require(isinstance(rows, list), f"{path.name}: {key} rows missing")
    return rows


def verify_chronology(
    intervals: dict[str, tuple[float, float]],
    chronology: dict[str, Any],
    cleanup: dict[str, Any] | None,
) -> None:
    """Validate receipt ordering against an immutable post-decision witness.

    Artifact mtimes are deliberately read from chronology.json.  A fresh
    checkout may rewrite all packet mtimes, so comparing them live would turn
    a valid committed packet into a false chronology failure.
    """
    require(chronology.get("schema") == "litchi.performance.0815.chronology.v1",
            "chronology schema changed")
    captured_at = chronology.get("captured_at")
    require(isinstance(captured_at, (int, float)) and not isinstance(captured_at, bool)
            and math.isfinite(captured_at) and captured_at > 0,
            "chronology capture time is invalid")
    files = chronology.get("files")
    expected_paths = {
        "application": P / "application.json",
        "qualification-audit": P / "qualification-audit.json",
        "analysis": P / "analysis.json",
        "root-audit": P / "root-audit.json",
        "profile-analysis": P / "profile-analysis.json",
        "decision": P / "decision.json",
        "disposition": P / "disposition.json",
    }
    require(isinstance(files, dict) and set(files) == set(expected_paths),
            "chronology artifact set changed")
    observed: dict[str, float] = {}
    for name, expected in expected_paths.items():
        value = files.get(name)
        require(isinstance(value, dict), f"chronology {name} witness is malformed")
        witnessed = packet_path(value.get("path"), f"chronology {name}")
        require(witnessed == expected,
                f"chronology {name} path changed")
        require(type(value.get("bytes")) is int and value["bytes"] >= 0,
                f"chronology {name} byte count is invalid")
        require(is_sha(value.get("sha256")),
                f"chronology {name} SHA-256 is invalid")
        require(witnessed.is_file() and not witnessed.is_symlink()
                and witnessed.stat().st_size == value["bytes"]
                and sha(witnessed) == value["sha256"],
                f"chronology {name} artifact identity changed")
        observed_ns = value.get("observed_mtime_ns")
        require(type(observed_ns) is int and observed_ns > 0,
                f"chronology {name} observed mtime is invalid")
        observed[name] = observed_ns / 1e9
        require(captured_at >= observed[name],
                f"chronology capture precedes {name} observation")

    def after(left: str, right: str) -> None:
        require(intervals[left][1] <= intervals[right][0],
                f"chronology overlap: {left} -> {right}")

    after("build-before", "probe-before")
    after("probe-before", "qualification")
    require(observed["application"] >= observed["qualification-audit"],
            "application precedes qualification audit")
    require(observed["application"] >= intervals["qualification"][1],
            "application precedes qualification completion")
    require(observed["application"] <= intervals["quality-after"][0],
            "after quality precedes candidate application")
    after("quality-after", "build-after")
    after("build-after", "probe-after")
    after("probe-after", "native")
    after("native", "allocation")
    after("allocation", "profiles")
    capture_end = max(intervals["native"][1], intervals["allocation"][1], intervals["profiles"][1])
    for name in ("analysis", "root-audit", "profile-analysis"):
        require(observed[name] >= capture_end,
                f"{name} precedes terminal capture")
    decision_time = observed["decision"]
    for name in ("analysis", "root-audit", "profile-analysis"):
        require(decision_time >= observed[name],
                f"decision precedes {name}")
    require(observed["disposition"] >= decision_time,
            "disposition precedes decision")
    if cleanup is not None:
        started, ended = cleanup.get("started"), cleanup.get("ended")
        require(isinstance(started, (int, float)) and not isinstance(started, bool)
                and math.isfinite(started)
                and isinstance(ended, (int, float)) and not isinstance(ended, bool)
                and math.isfinite(ended) and started > observed["disposition"]
                and ended >= started,
                "cleanup interval does not follow disposition witness")


def run_reader_checks() -> dict[str, str]:
    outputs = {}
    for script, output in READER_CHECKS:
        require((P / output).is_file(), f"reader output is missing: {output}")
        result = subprocess.run(
            [sys.executable, "-B", str(P / script), "--check"],
            cwd=ROOT,
            text=True,
            capture_output=True,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        )
        require(result.returncode == 0,
                f"{script} --check failed: {result.stdout}{result.stderr}")
        outputs[script] = (result.stdout + result.stderr).strip()
    qualification = P / "qualification-audit.json"
    require(qualification.is_file(), "reader output is missing: qualification-audit.json")
    result = subprocess.run(
        [sys.executable, "-B", str(P / "analysis.py"), "--qualification", "--check"],
        cwd=ROOT,
        text=True,
        capture_output=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
    )
    require(result.returncode == 0,
            f"analysis.py --qualification --check failed: {result.stdout}{result.stderr}")
    outputs["analysis.py --qualification"] = (result.stdout + result.stderr).strip()
    return outputs


def validate(final: bool) -> dict[str, Any]:
    custody = verify_origin()
    verify_frozen_build_inputs()
    plan = verify_plan()
    cleanup = read(P / "cleanup.json") if (P / "cleanup.json").is_file() else None
    builds, before, after = verify_builds(cleanup)
    expected_after = candidate_sources(before)
    require(after == expected_after, "after build differs from candidate archive")
    verify_codegen(builds)
    codegen_analysis = verify_codegen_analysis(builds)
    verify_quality_reuse(before, custody["root_inputs"])
    verify_lane("qualification", plan, builds, before, cleanup)
    verify_qualification_audit(before)
    verify_application(before, after)
    native_start = interval(
        timestamp_rows(P / "native/receipts.json", "self"), "native"
    )[0]
    verify_codegen_gate(codegen_analysis, native_start)
    native_interval = verify_lane("native", plan, builds, after, cleanup)
    verify_lane("allocation", plan, builds, after, cleanup)
    intervals = {
        "build-before": interval(timestamp_rows(P / "build-before/build.json", "rows"), "build-before"),
        "probe-before": interval(timestamp_rows(P / "probe-quality-before/receipts.json", "self"), "probe-before"),
        "qualification": interval(timestamp_rows(P / "qualification/receipts.json", "self"), "qualification"),
        "quality-after": interval(timestamp_rows(P / "quality-after/checks.json", "self"), "quality-after"),
        "build-after": interval(timestamp_rows(P / "build-after/build.json", "rows"), "build-after"),
        "probe-after": interval(timestamp_rows(P / "probe-quality-after/receipts.json", "self"), "probe-after"),
        "native": interval(timestamp_rows(P / "native/receipts.json", "self"), "native"),
        "allocation": interval(timestamp_rows(P / "allocation/receipts.json", "self"), "allocation"),
        "profiles": interval(timestamp_rows(P / "profiles/receipts.json", "self"), "profiles"),
    }
    quality = read(P / "quality-summary.json")
    verify_quality_summary(after, custody["root_inputs"])
    numeric = verify_reader_agreement()
    disposition = verify_decision_and_disposition(before, after, numeric)
    cleanup = verify_cleanup(final, builds, disposition)
    chronology = read(P / "chronology.json")
    verify_chronology(intervals, chronology, cleanup)
    reader_outputs = run_reader_checks()
    return {
        "schema": "litchi.performance.0815.validation.v1",
        "final": final,
        "counts": {"reports": 310, "samples": 6718, "main_reports": 306,
                    "main_samples": 6714, "profile_reports": 4, "profile_samples": 4},
        "source": {"before_files": len(before["files"]), "after_files": len(after["files"]),
                    "allowlist": sorted(SOURCE_ALLOWLIST),
                    "retained": disposition["production_change_retained"]},
        "policy": {
            "adoption_eligible": numeric["analysis"]["decision_guards"]["adoption_eligible"],
            "main_root_agree": True,
        },
        "quality_summary_schema": quality["schema"],
        "cleanup": {"present": cleanup is not None, "required": final},
        "reader_checks": {name: "PASS" for name in reader_outputs},
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", action="store_true",
                        help="require the post-disposition six-binary cleanup witness")
    args = parser.parse_args()
    result = validate(args.final)
    print(json.dumps(result, indent=2, sort_keys=True))
    print("0815 aggregate validation PASS", flush=True)


if __name__ == "__main__":
    try:
        main()
    except AssertionError as error:
        print(f"0815 aggregate validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
