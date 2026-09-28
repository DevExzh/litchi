"""Offline custody validator for the terminal 0808 early-stop packet.

The packet stopped after the candidate's production quality gate four.  This
reader intentionally has no timing or profiler path: it validates the
18-report before-only qualification, the before probe quality receipt, the
four-row after-quality stop, the independent baseline control, and the exact
source restoration.  It is valid both while the owned target is live and
after ``cleanup_early_stop.py`` has removed it.
"""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0808")
BASE_REVISION = "d28e3dc702d84f8752a2e329c2d3fc15f5c94f06"
SOURCE_ALLOWLIST = {"crates/litchi-pptx/src/notes/codec.rs"}
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple((shape, mode) for shape in SHAPES for mode in MODES)
DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100),
              "vendor": (12, 8), "unicode-vendor": (12, 8), "valid-4attr": (12, 8)}
TIMING = {
    "capture": "Package::opened_presentation only",
    "commit": "Transaction::commit only; package capture and one set_shape_text staging are outside the clock",
    "lifecycle": "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes",
}
PROBE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROBE_TOOL = "public-pptx-probe-0806"
MARKER = "litchi-perf-0780-static-mce-capabilities"
VALID_URIS = ["urn:litchi:perf:0806:extension:one", "urn:litchi:perf:0806:extension:two",
              "urn:litchi:perf:0806:extension:three", "urn:litchi:perf:0806:extension:four"]
VALID_NAMES = ["lx1:probeOne", "lx2:probeTwo", "lx3:probeThree", "lx4:probeFour"]
VALID_VALUES = ["litchi-perf-0806-valid-4attr-one", "litchi-perf-0806-valid-4attr-two",
                "litchi-perf-0806-valid-4attr-three", "litchi-perf-0806-valid-4attr-four"]
VENDOR_URIS = [
    "Xttp://schemas.openxmlformats.org/presentationml/2006/main",
    "Xttp://purl.oclc.org/ooxml/presentationml/main",
    "Xttp://schemas.openxmlformats.org/drawingml/2006/main",
    "Xttp://purl.oclc.org/ooxml/drawingml/main",
    "Xttp://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "Xttp://purl.oclc.org/ooxml/officeDocument/relationships",
]
VENDOR_NAMES = ["vP:probeP", "vPS:probePS", "vA:probeA", "vAS:probeAS",
                "vR:probeR", "vRS:probeRS"]
RAW_ALLOCATION = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
RESULT = re.compile(
    r"test result: (?:ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
    r"(\d+) ignored;\s+(\d+) measured;"
)
ERROR = re.compile(
    r"error: called `\.err\(\)\.expect\(\)` on a `Result` value\s*"
    r"--> crates/litchi-pptx/src/opened/tests\.rs:(\d+):(\d+)"
)


class EarlyStopError(RuntimeError):
    pass


def require(value: bool, message: str) -> None:
    if not value:
        raise EarlyStopError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EarlyStopError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def integer(value: Any, label: str, *, positive: bool = False) -> None:
    require(isinstance(value, int) and not isinstance(value, bool)
            and (value > 0 if positive else value >= 0), f"{label}: invalid integer")


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path missing")
    path = (Path(value) if Path(value).is_absolute() else P / value).resolve()
    try:
        path.relative_to(P.resolve())
    except ValueError as error:
        raise EarlyStopError(f"{label}: path escapes packet") from error
    return path


def packet_artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact malformed")
    path = packet_path(value.get("path"), label)
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 malformed")
    require(path.is_file() and not path.is_symlink(), f"{label}: file missing")
    require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
            f"{label}: identity changed")
    return path


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    return source_value(read(path), label)


def source_value(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            f"{label}: source manifest malformed")
    files = value["files"]
    require(len(files) == 9196 and all(isinstance(name, str) and is_sha(digest)
                                       for name, digest in files.items()),
            f"{label}: source census malformed")
    return {"revision": value.get("revision"), "files": dict(files)}


def current_source() -> dict[str, str]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", "crates", "Cargo.toml", "clippy.toml",
         ".cargo/config.toml", "rust-toolchain.toml"], cwd=ROOT)
    names = [name for name in raw.decode().split("\0") if name]
    return {name: sha(ROOT / name) for name in names}


def build_before() -> dict[str, Any]:
    manifest = read(P / "build-before/build.json")
    require(manifest.get("schema") == "litchi.performance.0808.build-before.v1",
            "before build schema changed")
    source = source_manifest(packet_artifact(manifest.get("source"), "before build source"),
                             "before build source")
    require(source["revision"] == BASE_REVISION and current_source() == source["files"],
            "restored production source differs from before build")
    require(manifest.get("probe") == {
        name: sha(P / name) for name in sorted(manifest.get("probe", {}))
    }, "before probe manifest changed")
    for name in manifest["probe"]:
        path = packet_path(name, f"before probe {name}")
        require(path.is_file() and sha(path) == manifest["probe"][name],
                f"before probe source changed: {name}")
    rows = manifest.get("rows")
    require(isinstance(rows, list) and len(rows) == 3
            and {row.get("name") for row in rows} == {"native", "allocation", "profile"}
            and all(row.get("exit_code") == 0 for row in rows), "before build rows changed")
    for row in rows:
        packet_artifact(row.get("log"), f"before {row['name']} build log")
    binaries = manifest.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
            "before binary matrix changed")
    normalized = {}
    for kind, value in binaries.items():
        require(isinstance(value, dict) and Path(value.get("path", "")).is_absolute(),
                f"before {kind} binary malformed")
        path = Path(value["path"])
        require(path.parent == TARGET and path.name == f"before-{kind}",
                f"before {kind} binary path changed")
        normalized[kind] = external_binary(value, f"before {kind}")
    return {"manifest": manifest, "source": source, "binaries": binaries,
            "binary_digests": normalized}


def cleanup_witness() -> dict[str, Any] | None:
    path = P / "early-stop-cleanup.json"
    if not path.exists():
        require(TARGET.is_dir(), "owned target missing without cleanup witness")
        return None
    value = read(path)
    require(value.get("schema") == "litchi.performance.0808.early-stop-cleanup.v1"
            and value.get("status") == "stopped-before-after-build"
            and value.get("target") == str(TARGET)
            and value.get("target_removed") is True
            and value.get("verified_before_removal") is True
            and not TARGET.exists(), "early-stop cleanup witness changed")
    removed = value.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 3
            and {row.get("kind") for row in removed} == {"native", "allocation", "profile"},
            "early-stop cleanup binary set changed")
    return value


def external_binary(value: dict[str, Any], label: str) -> dict[str, Any]:
    path = Path(value.get("path", ""))
    require(path.is_absolute() and path.parent == TARGET and isinstance(value.get("bytes"), int)
            and value["bytes"] > 0 and is_sha(value.get("sha256")),
            f"{label}: binary identity malformed")
    if path.is_file():
        require(not path.is_symlink() and path.stat().st_size == value["bytes"]
                and sha(path) == value["sha256"], f"{label}: live binary changed")
    else:
        cleanup = read(P / "early-stop-cleanup.json")
        rows = cleanup.get("removed_binaries", [])
        matches = [row for row in rows if isinstance(row, dict) and row.get("path") == str(path)]
        require(len(matches) == 1 and matches[0].get("bytes") == value["bytes"]
                and matches[0].get("sha256") == value["sha256"],
                f"{label}: cleanup witness does not identify binary")
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}


def candidate_chain(before: dict[str, Any]) -> None:
    application = read(P / "application.json")
    require(application.get("schema") == "litchi.performance.0808.application.v1"
            and application.get("allowlist") == sorted(SOURCE_ALLOWLIST),
            "candidate application schema changed")
    source_before = source_value(application.get("source_before"), "application source before")
    candidate_source = source_value(application.get("source"), "application candidate source")
    require(source_before == before["source"], "application before source differs")
    changed = {name for name in before["source"]["files"]
               if before["source"]["files"].get(name) != candidate_source["files"].get(name)}
    require(changed == SOURCE_ALLOWLIST and candidate_source["revision"] == BASE_REVISION,
            "application candidate scope changed")
    packet_artifact(application.get("manifest"), "application manifest")
    packet_artifact(application.get("patch"), "application patch")
    manifest = read(P / "candidate/manifest.json")
    require(manifest.get("schema") == "litchi.performance.0808.candidate-manifest.v1"
            and manifest.get("base_commit") == BASE_REVISION
            and manifest.get("production_source_changed") is False,
            "candidate manifest changed")
    rows = list(manifest.get("files", {}).values())
    require(len(rows) == 1 and {row.get("production_path") for row in rows} == SOURCE_ALLOWLIST,
            "candidate manifest file scope changed")
    for row in rows:
        for side, expected in (("before", before["source"]["files"][row["production_path"]]),
                               ("after", candidate_source["files"][row["production_path"]])):
            path = packet_artifact(row[side], f"candidate {side} source")
            require(sha(path) == expected, f"candidate {side} source differs")
    restored = source_manifest(packet_artifact(read(P / "disposition.json").get("restored_source"),
                                                "restored source"), "restored source")
    require(restored == before["source"], "restored source differs from before")
    disposition = read(P / "disposition.json")
    require(disposition.get("schema") == "litchi.performance.0808.disposition.v1"
            and disposition.get("status") == "rejected"
            and disposition.get("production_change_retained") is False,
            "rejected disposition missing")


def verify_fixture(report: dict[str, Any], shape: str, label: str) -> None:
    require(report.get("schema") == PROBE_SCHEMA and report.get("tool") == PROBE_TOOL
            and report.get("marker") == MARKER and report.get("shape") == shape
            and (report.get("slides"), report.get("shapes_per_slide")) == DIMENSIONS[shape]
            and report.get("timing_scope") in TIMING.values(), f"{label}: probe identity changed")
    fixture = report.get("fixture")
    require(isinstance(fixture, dict), f"{label}: fixture missing")
    if shape == "valid-4attr":
        require(fixture == {
            "injection": "valid-four-distinct-namespaced-extension-attributes",
            "slide_parts": 12, "replaced_text_tags": 96, "namespace_declarations": 4,
            "namespaced_attributes": 4, "namespace_uris": VALID_URIS,
            "attribute_names": VALID_NAMES,
        }, f"{label}: valid fixture changed")
    elif shape in ("vendor", "unicode-vendor"):
        injection = "same-length-known-uri-near-misses" if shape == "vendor" else \
            "same-length-valid-utf8-unknown-uris"
        require(fixture.get("injection") == injection and fixture.get("slide_parts") == 12
                and fixture.get("replaced_text_tags") == 96
                and fixture.get("namespace_declarations") == 6
                and fixture.get("namespaced_attributes") == 6
                and fixture.get("namespace_uris") == (VENDOR_URIS if shape == "vendor" else [
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu",
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu",
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu",
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuu",
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu",
                    "urn:vendor:éuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuuu" ])
                and fixture.get("attribute_names") == VENDOR_NAMES,
                f"{label}: vendor fixture changed")
    else:
        require(fixture == {"injection": "none", "slide_parts": 0,
                            "replaced_text_tags": 0, "namespace_declarations": 0,
                            "namespaced_attributes": 0, "namespace_uris": [],
                            "attribute_names": []}, f"{label}: ordinary fixture changed")


def verify_sample(report: dict[str, Any], historical: dict[str, Any], shape: str,
                  mode: str, label: str) -> None:
    verify_fixture(report, shape, label)
    require(report.get("warmup") == 0 and report.get("samples_requested") == 1,
            f"{label}: qualification policy changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{label}: sample count changed")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("index") == 0
            and isinstance(sample.get("elapsed_ns"), int) and sample["elapsed_ns"] > 0,
            f"{label}: elapsed sample changed")
    metrics = sample.get("metrics")
    require(isinstance(metrics, dict) and metrics.get("elapsed_ns") == sample["elapsed_ns"]
            and metrics.get("slides") == report["slides"]
            and metrics.get("shapes_per_slide") == report["shapes_per_slide"],
            f"{label}: metrics changed")
    if mode == "capture":
        require(metrics.get("captured_slides") == report["slides"]
                and metrics.get("captured_shapes_per_slide") == report["shapes_per_slide"],
                f"{label}: capture metrics changed")
    else:
        require("captured_slides" not in metrics and "captured_shapes_per_slide" not in metrics,
                f"{label}: non-capture metrics changed")
    source = report.get("source")
    output = sample.get("output")
    verification = sample.get("verification")
    allocation = sample.get("allocation")
    require(isinstance(source, dict) and is_sha(source.get("sha256"))
            and isinstance(output, dict) and is_sha(output.get("sha256"))
            and isinstance(output.get("bytes"), int) and output["bytes"] > 0
            and isinstance(verification, dict) and isinstance(allocation, dict),
            f"{label}: raw oracle objects malformed")
    require(sample.get("source_sha256") == source["sha256"]
            and verification.get("readback_bytes") == output["bytes"]
            and verification.get("readback_sha256") == output["sha256"]
            and verification.get("semantic_check") is True
            and verification.get("reopened") is True
            and verification.get("expected_text") == verification.get("actual_text"),
            f"{label}: semantic oracle changed")
    require(all(isinstance(allocation.get(field), int) and allocation[field] >= 0
                for field in RAW_ALLOCATION)
            and allocation.get("status") == "measured"
            and allocation.get("scope") == "operation_global_system_allocator"
            and allocation["live_bytes_after"] == allocation["live_bytes_before"]
            + allocation["allocated_bytes"] - allocation["deallocated_bytes"]
            and allocation["failed_allocation_calls"] == 0
            and allocation["region_peak_live_bytes"] >= allocation["live_bytes_before"]
            and allocation["region_peak_live_bytes"] >= allocation["live_bytes_after"]
            and allocation["peak_live_bytes_after"] >= allocation["region_peak_live_bytes"],
            f"{label}: allocation oracle changed")
    if shape == "valid-4attr":
        require(verification.get("extension_preservation_check") is True
                and verification.get("extension_text_tags") == 96
                and verification.get("extension_attributes_per_text_tag") == 4
                and verification.get("extension_attribute_occurrences") == 384
                and verification.get("extension_value_occurrences") == 384
                and verification.get("extension_namespace_declarations_per_slide") == 4
                and verification.get("extension_namespace_uris") == VALID_URIS
                and verification.get("extension_attribute_names") == VALID_NAMES
                and verification.get("extension_attribute_values") == VALID_VALUES,
                f"{label}: extension oracle changed")
    else:
        require(all(verification.get(field) is None for field in (
            "extension_preservation_check", "extension_text_tags",
            "extension_attributes_per_text_tag", "extension_attribute_occurrences",
            "extension_value_occurrences", "extension_namespace_declarations_per_slide",
            "extension_namespace_uris", "extension_attribute_names", "extension_attribute_values")),
                f"{label}: unexpected extension oracle")
    vendor = shape in ("vendor", "unicode-vendor")
    require(verification.get("unknown_namespace_check") is (True if vendor else None)
            and verification.get("unknown_namespace_occurrences") == (96 if vendor else None)
            and verification.get("marker_matches") is (True if mode in ("commit", "lifecycle") else None),
            f"{label}: namespace or marker oracle changed")
    # The report is compared to the sealed 0806 raw qualification after
    # removing only measured elapsed values.  Output, allocation, and the
    # complete verification map therefore remain exact historical oracles.
    current = json.loads(json.dumps(sample, sort_keys=True))
    prior = json.loads(json.dumps(historical["samples"][0], sort_keys=True))
    current.pop("elapsed_ns", None)
    prior.pop("elapsed_ns", None)
    current.get("metrics", {}).pop("elapsed_ns", None)
    prior.get("metrics", {}).pop("elapsed_ns", None)
    require(current == prior, f"{label}: historical non-timing oracle changed")


def qualification(before: dict[str, Any]) -> dict[str, Any]:
    complete = read(P / "qualification/complete.json")
    require(complete.get("schema") == "litchi.performance.0808.qualification.complete.v1"
            and complete.get("children") == 18 and complete.get("reports") == 18
            and complete.get("samples") == 18, "qualification completion changed")
    source = source_manifest(packet_artifact(complete.get("source"), "qualification source"),
                             "qualification source")
    require(source == before["source"], "qualification source differs from before")
    receipts = read(packet_artifact(complete.get("receipts"), "qualification receipts"))
    require(isinstance(receipts, list) and len(receipts) == 18, "qualification receipt count changed")
    historical_dir = ROOT / "docs/performance/results/change-0806/qualification"
    rows = []
    for receipt, (shape, mode) in zip(receipts, CASES):
        label = f"qualification/{shape}/{mode}"
        require(receipt.get("schema") == "litchi.performance.0808.capture-receipt.v1"
                and receipt.get("lane") == "qualification" and receipt.get("block") == 0
                and receipt.get("shape") == shape and receipt.get("mode") == mode
                and receipt.get("leg") == "before" and receipt.get("exit_code") == 0,
                f"{label}: receipt identity changed")
        require(receipt.get("binary") == before["binaries"]["allocation"],
                f"{label}: allocation binary changed")
        report_path = packet_artifact(receipt.get("report"), f"{label} report")
        packet_artifact(receipt.get("log"), f"{label} log")
        rss = packet_artifact(receipt.get("rss"), f"{label} RSS")
        require(rss.read_text(encoding="utf-8").strip().isdigit()
                and int(rss.read_text(encoding="utf-8").strip()) > 0, f"{label}: RSS changed")
        historical_path = historical_dir / f"0-{shape}-{mode}-before.json"
        historical = read(historical_path)
        report = read(report_path)
        require(report.get("schema") == historical.get("schema")
                and report.get("tool") == historical.get("tool")
                and report.get("mode") == historical.get("mode")
                and report.get("shape") == historical.get("shape")
                and report.get("slides") == historical.get("slides")
                and report.get("shapes_per_slide") == historical.get("shapes_per_slide")
                and report.get("timing_scope") == historical.get("timing_scope")
                and report.get("marker") == historical.get("marker")
                and report.get("source") == historical.get("source")
                and report.get("fixture") == historical.get("fixture")
                and report.get("allocator") == historical.get("allocator"),
                f"{label}: historical report identity changed")
        verify_sample(report, historical, shape, mode, label)
        rows.append({"shape": shape, "mode": mode, "report": str(report_path.relative_to(P)),
                     "report_sha256": sha(report_path)})
    return {"reports": 18, "samples": 18, "rows": rows, "timings_imported": False}


def probe_quality(before: dict[str, Any]) -> dict[str, Any]:
    complete = read(P / "probe-quality-before/complete.json")
    require(complete.get("schema") == "litchi.performance.0808.probe-quality-before.v1"
            and complete.get("gate_count") == 3 and complete.get("tests_passed") == 36,
            "before probe quality completion changed")
    inputs = read(packet_artifact(complete.get("inputs"), "probe quality inputs"))
    require(inputs.get("source", {}).get("files") == before["source"]["files"],
            "before probe quality source changed")
    probe = inputs.get("probe")
    require(isinstance(probe, dict), "before probe quality probe manifest missing")
    for name, digest in probe.items():
        path = packet_path("probe-src/" + name if not name.startswith("probe-src/") else name,
                           f"probe quality input {name}")
        require(is_sha(digest) and sha(path) == digest, f"probe quality input changed: {name}")
    receipts = read(packet_artifact(complete.get("receipts"), "probe quality receipts"))
    require(isinstance(receipts, list) and len(receipts) == 3, "probe quality receipt count changed")
    for index, row in enumerate(receipts, 1):
        expected = [
            ["cargo", "fmt", "--manifest-path", str(P / "probe-src/Cargo.toml"), "--", "--check"],
            ["cargo", "test", "--offline", "--locked", "--release", "--manifest-path",
             str(P / "probe-src/Cargo.toml"), "--all-features", "--", "--test-threads=1"],
            ["cargo", "clippy", "--offline", "--locked", "--release", "--manifest-path",
             str(P / "probe-src/Cargo.toml"), "--all-features", "--all-targets", "--", "-D", "warnings"],
        ][index - 1]
        require(row.get("gate") == index and row.get("exit_code") == 0
                and row.get("command") == expected, f"probe quality gate {index} failed")
        log = packet_artifact(row.get("log"), f"probe quality gate {index} log")
        if index == 2:
            matches = list(RESULT.finditer(log.read_text(encoding="utf-8", errors="replace")))
            require(len(matches) == 1 and tuple(map(int, matches[0].groups())) == (36, 0, 0, 0),
                    "probe quality test count changed")
    return {"gates": 3, "tests_passed": 36,
            "receipts": [{"gate": row["gate"], "log": str(Path(row["log"]["path"]).relative_to(P)),
                          "log_sha256": row["log"]["sha256"]} for row in receipts]}


def production_quality_stop(application: dict[str, Any]) -> dict[str, Any]:
    source = source_manifest(P / "quality-after/source.json", "after quality source")
    require(source == source_value(application.get("source"), "application candidate source"),
            "after quality source differs from application")
    checks = read(P / "quality-after/checks.json")
    require(isinstance(checks, list) and len(checks) == 4
            and [row.get("exit_code") for row in checks] == [0, 0, 0, 101],
            "after quality did not stop at gate four")
    for index, row in enumerate(checks[:3], 1):
        expected = [
            ["cargo", "fmt", "-p", "litchi-pptx", "--", "--check"],
            ["cargo", "check", "--offline", "--locked", "-p", "litchi-pptx",
             "--all-features", "--all-targets"],
            ["cargo", "test", "--offline", "--locked", "-p", "litchi-pptx",
             "--all-features", "--", "--test-threads=2"],
        ][index - 1]
        require(row.get("gate") == index and row.get("command") == expected,
                f"after quality gate {index} receipt changed")
        packet_artifact(row.get("log"), f"after quality gate {index} log")
    failed = checks[3]
    require(failed.get("gate") == 4 and failed.get("command") == [
        "cargo", "clippy", "--offline", "--locked", "-p", "litchi-pptx",
        "--all-features", "--all-targets", "--", "-D", "warnings"],
            "after quality failed command changed")
    failed_log = packet_artifact(failed.get("log"), "after quality Clippy log")
    text = failed_log.read_text(encoding="utf-8")
    errors = ERROR.findall(text)
    require(errors == [("464", "14"), ("538", "10"), ("557", "10")]
            and text.count("error: called `.err().expect()` on a `Result` value") == 3
            and "clippy::err-expect" in text and "clippy::err_expect" in text,
            "after quality diagnostics changed")
    test_log = packet_artifact(checks[2].get("log"), "after quality test log")
    suites = [tuple(map(int, match.groups())) for match in RESULT.finditer(
        test_log.read_text(encoding="utf-8", errors="replace"))]
    require(len(suites) == 85 and tuple(sum(row[index] for row in suites) for index in range(4))
            == (1241, 0, 3, 0), "after quality test totals changed")
    require(not (P / "quality-after.json").exists() and not (P / "quality-before.json").exists(),
            "production quality completion receipt must be absent")
    return {"gates_executed": 4, "passes": 3, "failures": 1, "failed_gate": 4,
            "test_groups": 85, "tests": {"passed": 1241, "failed": 0, "ignored": 3, "measured": 0},
            "failure_log": str(failed_log.relative_to(P)), "failure_log_sha256": sha(failed_log),
            "diagnostic_locations": ["464:14", "538:10", "557:10"]}


def baseline_control() -> dict[str, Any]:
    receipt = read(P / "baseline-clippy.json")
    require(receipt.get("schema") == "litchi.performance.0808.baseline-clippy.v1"
            and receipt.get("command") == [
                "cargo", "clippy", "--offline", "--locked", "-p", "litchi-pptx",
                "--all-features", "--all-targets", "--", "-D", "warnings"]
            and receipt.get("exit_code") == 101
            and receipt.get("restored_source_sha256") == sha(P / "restored-source.json"),
            "baseline Clippy control receipt changed")
    log = packet_artifact(receipt.get("log"), "baseline Clippy control log")
    text = log.read_text(encoding="utf-8")
    errors = ERROR.findall(text)
    require(errors == [("464", "14"), ("538", "10"), ("557", "10")]
            and text.count("error: called `.err().expect()` on a `Result` value") == 3
            and "clippy::err-expect" in text and "clippy::err_expect" in text,
            "baseline control diagnostics changed")
    return {"exit_code": 101, "log": str(log.relative_to(P)), "log_sha256": sha(log),
            "diagnostic_locations": ["464:14", "538:10", "557:10"]}


def absence() -> None:
    for name in ("build-after", "native", "allocation", "profiles", "profile-analysis.json",
                 "quality-before.json", "quality-after.json", "analysis.json", "root-audit.json"):
        require(not (P / name).exists(), f"unstarted or completed evidence unexpectedly exists: {name}")


def validate(require_cleanup: bool) -> dict[str, Any]:
    frozen = read(P / "build-before/frozen-inputs.json")
    require(set(frozen) == {
        "plan.json", "adoption-policy.json", "analysis-plan.json", "custody.py",
        "build.py", "capture.py", "profile.py", "quality.py", "probe_quality.py",
        "apply_candidate.py", "restore_candidate.py", "origin.json", "host.json",
        "inheritance.json", "architecture-inputs.json",
    }, "frozen input census changed")
    for name, digest in frozen.items():
        require(sha(P / name) == digest, f"frozen input changed: {name}")
    architecture = read(P / "architecture-inputs.json")
    require(len(architecture) == 35, "architecture input census changed")
    for name, digest in architecture.items():
        require(sha(ROOT / name) == digest, f"architecture input changed: {name}")
    for name, digest in read(P / "origin.json")["unrelated"].items():
        require(sha(ROOT / name) == digest, f"unrelated file changed: {name}")
    decision = read(P / "decision.json")
    require(decision.get("adoption_eligible") is False
            and decision.get("production_adoption") is False
            and decision.get("disposition") == "deferred-before-performance",
            "early-stop decision changed")
    accepted = read(P / "qualification-audit.json")
    require(accepted.get("passed") is True and accepted.get("accepted_before_application") is True,
            "before-only qualification acceptance missing")
    subprocess.run([sys.executable, "-B", str(P / "analysis.py"),
                    "--qualification", "--check"], cwd=ROOT,
                   stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    plan = read(P / "early-stop-plan.json")
    require(plan.get("schema") == "litchi.performance.0808.early-stop-plan.v1"
            and plan.get("status") == "stopped-before-after-build", "early-stop plan changed")
    before = build_before()
    cleanup = cleanup_witness()
    candidate_chain(before)
    application = read(P / "application.json")
    qualification_result = qualification(before)
    probe_result = probe_quality(before)
    quality_result = production_quality_stop(application)
    control_result = baseline_control()
    absence()
    if require_cleanup:
        require(cleanup is not None, "early-stop cleanup witness required")
    return {
        "schema": "litchi.performance.0808.early-stop-validation.v1",
        "status": "stopped-before-after-build",
        "decision": "deferred-before-performance",
        "source": {"before_revision": BASE_REVISION, "file_count": 9196,
                    "restored_current": True, "allowlist": sorted(SOURCE_ALLOWLIST)},
        "qualification": qualification_result,
        "probe_quality_before": probe_result,
        "production_quality_after": quality_result,
        "baseline_control": control_result,
        "binaries": {kind: before["binary_digests"][kind] for kind in ("native", "allocation", "profile")},
        "planned_work_not_started": ["build-after", "native", "allocation", "profiles"],
        "cleanup_contract": {"target": str(TARGET), "binary_count": 3,
                              "witness": "early-stop-cleanup.json"},
        "performance_measured": False,
        "adoption_eligible": False,
        "timings_imported": False,
    }


def main() -> None:
    args = set(sys.argv[1:])
    require(args <= {"--write", "--check", "--require-cleanup"}
            and ("--write" in args) ^ ("--check" in args),
            "use --write or --check, with optional --require-cleanup")
    value = validate("--require-cleanup" in args)
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    path = P / "early-stop-validation.json"
    if "--write" in args:
        path.write_text(encoded, encoding="utf-8")
    else:
        require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
                "early-stop-validation.json does not replay byte-for-byte")
    print("0808 early-stop validation PASS: 18 qualification reports, 3 probe gates, gate-4 stop")


if __name__ == "__main__":
    try:
        main()
    except (EarlyStopError, subprocess.CalledProcessError) as error:
        print(f"early-stop validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
