#!/usr/bin/env python3
"""Portable, fail-closed replay for the change-0431 matched comparison.

The replay verifies the frozen lane order, role-specific build and source
custody, every receipt and raw-or-gzip artifact binding, and every report with
the copied report verifier.  It then recomputes ``comparison.json`` through
the bound comparison driver.  No captured executable or repository checkout
is required for an exported bundle.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any

import compare


ROOT = Path(__file__).resolve().parent
CHANGE = 431
SAMPLES = 30
WARMUPS = 3
PROCESSES = 16
HEX40 = set("0123456789abcdef")


class Invalid(ValueError):
    """The retained bundle is not a valid 0431 replay."""


def fail(message: str) -> None:
    raise Invalid(message)


def digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def load_bytes(raw: bytes, label: str) -> Any:
    try:
        return json.loads(
            raw.decode("utf-8"),
            object_pairs_hook=compare.duplicate_pairs,
            parse_constant=compare.reject_constant,
        )
    except compare.Invalid:
        raise
    except (UnicodeError, json.JSONDecodeError) as exc:
        fail(f"{label}: invalid JSON: {exc}")


def safe_path(root: Path, name: str) -> Path:
    return compare.safe_path(root, name)


def logical_bytes(root: Path, name: str) -> bytes:
    return compare.logical_bytes(root, name)


def load(root: Path, name: str) -> Any:
    return load_bytes(logical_bytes(root, name), name)


def obj(value: Any, label: str) -> dict[str, Any]:
    return compare.obj(value, label)


def array(value: Any, label: str) -> list[Any]:
    return compare.array(value, label)


def text(value: Any, label: str) -> str:
    return compare.text(value, label)


def uint(value: Any, label: str) -> int:
    return compare.uint(value, label)


def sha256_text(value: Any, label: str) -> str:
    return compare.sha256_text(value, label)


def verify_protocol(root: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    protocol = obj(load(root, "protocol.json"), "protocol")
    expected_keys = {
        "change", "baseline_revision", "adr_tree", "rust_toolchain", "build_profile",
        "rustflags", "release_debug", "cpu", "samples_per_process", "warmups_per_process",
        "lanes", "order", "expected_processes_per_role", "regression_review_percent", "scope",
    }
    if set(protocol) != expected_keys:
        fail("protocol: fields differ from the frozen 0431 schema")
    if protocol["change"] != CHANGE:
        fail("protocol.change: expected 431")
    for key in ("baseline_revision", "adr_tree"):
        value = text(protocol[key], f"protocol.{key}")
        expected_length = 40 if key == "baseline_revision" else 40
        if len(value) != expected_length or any(char not in HEX40 for char in value):
            fail(f"protocol.{key}: expected lowercase revision identity")
    if protocol["rust_toolchain"] != "1.98.1" or protocol["build_profile"] != "release":
        fail("protocol: toolchain/profile changed")
    if protocol["rustflags"] != "-Cforce-frame-pointers=yes" or protocol["release_debug"] != 1:
        fail("protocol: release flags changed")
    if protocol["cpu"] != 2 or protocol["samples_per_process"] != SAMPLES or protocol["warmups_per_process"] != WARMUPS:
        fail("protocol: CPU or sample policy changed")
    if protocol["expected_processes_per_role"] != PROCESSES or protocol["regression_review_percent"] != 5:
        fail("protocol: process or review threshold changed")
    if not isinstance(protocol["scope"], str) or not protocol["scope"]:
        fail("protocol.scope: missing scope")
    lanes = array(protocol["lanes"], "protocol.lanes")
    expected_lanes: list[dict[str, Any]] = []
    for corpus in ("plain", "media-rich"):
        for label, provider, cap, delay in (
            ("bytes", "bytes", None, None),
            ("file", "file", None, None),
            ("short-range", "range", 4096, 0),
            ("delayed-range", "range", 65536, 200),
        ):
            expected_lanes.append({
                "command": "provider-lifecycle", "selector_key": "corpus", "selector": corpus,
                "provider_label": label, "provider": provider, "max_range": cap, "delay_us": delay,
            })
    if lanes != expected_lanes:
        fail("protocol.lanes: lane definitions differ from frozen matrix")
    order = array(protocol["order"], "protocol.order")
    expected_order = [dict(lane, repeat="R1") for lane in expected_lanes] + [
        dict(lane, repeat="R2") for lane in reversed(expected_lanes)
    ]
    if order != expected_order:
        fail("protocol.order: expected forward R1 and reverse R2 order")
    return protocol, order


def verify_source_manifest(root: Path, value: Any, label: str) -> dict[str, str]:
    manifest = obj(value, label)
    if set(manifest) != {"path", "sha256", "files"}:
        fail(f"{label}: source manifest fields differ")
    path = text(manifest["path"], f"{label}.path")
    expected_digest = sha256_text(manifest["sha256"], f"{label}.sha256")
    files = uint(manifest["files"], f"{label}.files")
    raw = logical_bytes(root, path)
    if digest(raw) != expected_digest:
        fail(f"{label}: retained manifest hash differs")
    parsed = load_bytes(raw, f"{label}.payload")
    if not isinstance(parsed, dict) or len(parsed) != files:
        fail(f"{label}: source file count differs")
    for source, source_digest in parsed.items():
        if not isinstance(source, str) or not isinstance(source_digest, str) or len(source_digest) != 64:
            fail(f"{label}: malformed source entry")
        if any(char not in "0123456789abcdef" for char in source_digest):
            fail(f"{label}: malformed source digest")
    return parsed


def verify_build(root: Path, protocol: dict[str, Any], role: str) -> dict[str, Any]:
    build = obj(load(root, f"build-{role}.json"), f"build-{role}")
    required = {
        "role", "revision", "baseline_revision", "source_manifest", "binary_sha256", "binary_bytes",
        "capture_binary", "original_binary", "protocol_sha256", "verifier_sha256", "scope",
    }
    if not required.issubset(build):
        fail(f"build-{role}: required identity field is missing")
    if build["role"] != role or build["baseline_revision"] != protocol["baseline_revision"]:
        fail(f"build-{role}: role or baseline revision differs")
    revision = text(build["revision"], f"build-{role}.revision")
    if len(revision) != 40 or any(char not in HEX40 for char in revision):
        fail(f"build-{role}.revision: expected lowercase source revision")
    binary = sha256_text(build["binary_sha256"], f"build-{role}.binary_sha256")
    binary_bytes = uint(build["binary_bytes"], f"build-{role}.binary_bytes")
    if binary_bytes == 0:
        fail(f"build-{role}.binary_bytes: empty binary")
    if build["protocol_sha256"] != digest((root / "protocol.json").read_bytes()):
        fail(f"build-{role}.protocol_sha256: protocol custody mismatch")
    if build["verifier_sha256"] != digest((root / "verify-report.py").read_bytes()):
        fail(f"build-{role}.verifier_sha256: bound report verifier mismatch")
    text(build["capture_binary"], f"build-{role}.capture_binary")
    text(build["original_binary"], f"build-{role}.original_binary")
    text(build["scope"], f"build-{role}.scope")
    source_entries = verify_source_manifest(root, build["source_manifest"], f"build-{role}.source_manifest")
    return {
        "role": role,
        "revision": revision,
        "binary_sha256": binary,
        "binary_bytes": binary_bytes,
        "capture_binary": build["capture_binary"],
        "protocol_sha256": build["protocol_sha256"],
        "source_manifest": build["source_manifest"],
        "source_entries": source_entries,
        "raw": build,
    }


def verify_harness_source_identity(before: dict[str, Any], after: dict[str, Any]) -> int:
    """Bind the source-oracle/timing harness to identical role inputs."""
    before_entries = before["source_entries"]
    after_entries = after["source_entries"]
    before_harness = {
        path: value for path, value in before_entries.items()
        if path.startswith("tools/perf-baseline/") and path.endswith((".rs", ".toml", ".lock"))
    }
    after_harness = {
        path: value for path, value in after_entries.items()
        if path.startswith("tools/perf-baseline/") and path.endswith((".rs", ".toml", ".lock"))
    }
    if not before_harness:
        fail("source custody: tools/perf-baseline harness manifest is empty")
    if set(before_harness) != set(after_harness):
        fail("source custody: tools/perf-baseline file set differs between roles")
    changed = sorted(path for path in before_harness if before_harness[path] != after_harness[path])
    if changed:
        fail("source custody: tools/perf-baseline timing/oracle files differ: " + ", ".join(changed))
    return len(before_harness)


def run_checked(argv: list[str], label: str) -> None:
    result = subprocess.run(argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode:
        detail = (result.stderr or result.stdout).decode("utf-8", "replace").strip()
        fail(f"{label}: bound command failed: {detail or result.returncode}")


def verify_bound_report(root: Path, report_name: str) -> None:
    report_path = safe_path(root, report_name)
    verifier = safe_path(root, "verify-report.py")
    if report_path.is_file():
        run_checked([sys.executable, "-B", str(verifier), str(report_path)], report_name)
        return
    # JSON reports are retained raw by the capture/seal protocol.  Keep this
    # fallback fail-closed for an exported bundle that compressed one anyway.
    raw = logical_bytes(root, report_name)
    with tempfile.NamedTemporaryFile(prefix="litchi-0431-report-", suffix=".json") as temporary:
        temporary.write(raw)
        temporary.flush()
        run_checked([sys.executable, "-B", str(verifier), temporary.name], report_name)


def semantic_report_identity(report: dict[str, Any]) -> dict[str, Any]:
    identity = compare.input_identity(report)
    config_fields = ("schema", "corpus", "provider", "provider_config", "configured_limits", "destination_configured_limits")
    identity.update({field: report.get(field) for field in config_fields})
    return identity


def verify_role_captures(
    root: Path,
    role: str,
    build: dict[str, Any],
    protocol: dict[str, Any],
    order: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    index_name = f"{role}-index.json"
    index = array(load(root, index_name), index_name)
    if len(index) != PROCESSES or len(set(index)) != PROCESSES:
        fail(f"{index_name}: expected sixteen unique receipt paths")
    reports: list[dict[str, Any]] = []
    lane_identity: dict[tuple[str, str], dict[str, Any]] = {}
    lane_output: dict[tuple[str, str], dict[str, Any]] = {}
    sample_count = 0
    for receipt_name, lane in zip(index, order, strict=True):
        receipt_path = text(receipt_name, f"{index_name}.path")
        expected_path = f"{role}/{lane['selector']}-{lane['provider_label']}-{lane['repeat'].lower()}-receipt.json"
        if receipt_path != expected_path:
            fail(f"{index_name}: receipt order/path differs at {receipt_path}")
        compare_receipt, report, _ = compare.verify_receipt(
            root, receipt_path, lane, role, build["raw"], protocol
        )
        report_name = f"{role}/{compare_receipt['name']}.json"
        verify_bound_report(root, report_name)
        compare.report_lane_identity(report, lane, build)
        identity = semantic_report_identity(report)
        lane_key = (lane["selector"], lane["provider_label"])
        previous = lane_identity.setdefault(lane_key, identity)
        if previous != identity:
            fail(f"{role}/{compare_receipt['name']}: repeated input/config identity differs")
        output = compare.output_identity(report)
        if lane_key in lane_output and lane_output[lane_key] != output:
            fail(f"{role}/{compare_receipt['name']}: same-role output identity differs")
        lane_output[lane_key] = output
        rows = array(report.get("samples_raw"), report_name + ".samples_raw")
        sample_count += len(rows)
        if len(rows) != SAMPLES:
            fail(f"{report_name}: sample count differs from protocol")
        reports.append(report)
    if sample_count != PROCESSES * SAMPLES:
        fail(f"{role}: retained sample count differs from sixteen thirty-sample reports")
    return reports, sample_count


def verify_cross_role(
    before: list[dict[str, Any]], after: list[dict[str, Any]], order: list[dict[str, Any]]
) -> None:
    if len(before) != len(after) or len(before) != PROCESSES:
        fail("cross-role: report counts differ")
    for lane in order[:PROCESSES // 2]:
        key = (lane["selector"], lane["provider_label"])
        # Preserve the frozen order while avoiding any dependency on output
        # wrapper bytes.  Two same-lane reports are at positions 0..7 and
        # 8..15 in each role's index, with R2 reversed.
        lane_index = order.index(lane)
        before_r1 = before[lane_index]
        after_r1 = after[lane_index]
        # The report arrays are passed in index order; R2 index is the reverse
        # lane position in the second half.
        r2_index = PROCESSES - 1 - lane_index
        before_r2 = before[r2_index]
        after_r2 = after[r2_index]
        expected = compare.input_identity(before_r1)
        for label, report in (("before R2", before_r2), ("after R1", after_r1), ("after R2", after_r2)):
            if compare.input_identity(report) != expected:
                fail(f"cross-role {key}: {label} source/destination archive or manifest differs")
        if semantic_report_identity(before_r1) != semantic_report_identity(after_r1):
            fail(f"cross-role {key}: matched provider configuration differs")


def verify_compression(root: Path) -> int:
    path = root / "compression.json"
    if not path.is_file():
        return 0
    compression = obj(load(root, "compression.json"), "compression")
    count = 0
    for name, raw_row in compression.items():
        row = obj(raw_row, f"compression.{name}")
        if set(row) != {"original_path", "original_sha256", "original_bytes", "stored_sha256", "stored_bytes"}:
            fail(f"compression.{name}: fields differ")
        if not name.endswith(".gz") or row["original_path"] != name[:-3]:
            fail(f"compression.{name}: original path differs")
        stored = safe_path(root, name).read_bytes()
        raw = gzip.decompress(stored)
        if digest(stored) != sha256_text(row["stored_sha256"], f"compression.{name}.stored_sha256"):
            fail(f"compression.{name}: stored digest mismatch")
        if len(stored) != uint(row["stored_bytes"], f"compression.{name}.stored_bytes"):
            fail(f"compression.{name}: stored length mismatch")
        if digest(raw) != sha256_text(row["original_sha256"], f"compression.{name}.original_sha256"):
            fail(f"compression.{name}: raw digest mismatch")
        if len(raw) != uint(row["original_bytes"], f"compression.{name}.original_bytes"):
            fail(f"compression.{name}: raw length mismatch")
        count += 1
    return count


def verify_inventory(root: Path) -> int:
    path = root / "SHA256SUMS"
    if not path.is_file():
        return 0
    rows: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        if "  " not in line:
            fail("SHA256SUMS: malformed line")
        value, name = line.split("  ", 1)
        if name in rows or len(value) != 64 or any(char not in "0123456789abcdef" for char in value):
            fail(f"SHA256SUMS: malformed or duplicate entry {name}")
        target = safe_path(root, name)
        if not target.is_file() or digest(target.read_bytes()) != value:
            fail(f"SHA256SUMS: digest mismatch for {name}")
        rows[name] = value
    expected = {
        str(path.relative_to(root))
        for path in root.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    }
    if set(rows) != expected:
        fail("SHA256SUMS: inventory does not cover the bundle exactly")
    return len(rows)


def verify_checks(root: Path) -> int:
    checks = root / "checks"
    if not checks.is_dir():
        return 0
    count = 0
    for path in sorted(checks.glob("*.json")):
        row = obj(load(root, str(path.relative_to(root))), str(path))
        if row.get("change") != CHANGE:
            continue
        # Checks may also retain non-command custody/review artifacts (for
        # example a cleaned fuzz-source receipt).  Only check.py command
        # receipts have argv and receive command status/source/log policy.
        if "argv" not in row:
            continue
        if row.get("status") not in {"pass", "failed", "review"}:
            fail(f"{path}: invalid command receipt status")
        if row.get("driver_sha256") != digest((root / "check.py").read_bytes()):
            fail(f"{path}: command driver hash differs from check.py")
        if "source_before" in row or "source_after" in row:
            if row.get("source_before") != row.get("source_after"):
                fail(f"{path}: source custody changed during command")
            if "source_after" in row:
                verify_source_manifest(root, row["source_after"], f"{path}.source_after")
        log = row.get("log")
        if log is not None:
            log_obj = obj(log, f"{path}.log")
            log_name = text(log_obj.get("path"), f"{path}.log.path")
            raw = logical_bytes(root, log_name)
            if digest(raw) != sha256_text(log_obj.get("sha256"), f"{path}.log.sha256") or len(raw) != uint(log_obj.get("bytes"), f"{path}.log.bytes"):
                fail(f"{path}: command log custody mismatch")
        count += 1
    return count


def verify_planned_checks(root: Path) -> int:
    """Apply final required-pass policy without erasing historical failures."""
    path = root / "planned-checks.json"
    if not path.is_file():
        return 0
    planned = obj(load(root, "planned-checks.json"), "planned-checks")
    required_pass = planned.get("required_pass")
    if not isinstance(required_pass, list) or any(
        not isinstance(name, str) or not name or Path(name).name != name
        for name in required_pass
    ):
        fail("planned-checks.required_pass: expected unique receipt names")
    if len(set(required_pass)) != len(required_pass):
        fail("planned-checks.required_pass: duplicate receipt name")
    for name in required_pass:
        receipt = obj(load(root, f"checks/{name}.json"), f"planned check {name}")
        if receipt.get("status") != "pass":
            fail(f"planned check {name}: required pass receipt is not pass")
        if "source_unchanged" in receipt and receipt["source_unchanged"] is not True:
            fail(f"planned check {name}: source custody changed")
        if "exit_code" in receipt and (
            isinstance(receipt["exit_code"], bool) or receipt["exit_code"] != 0
        ):
            fail(f"planned check {name}: exit code is not zero")
    return len(required_pass)


def verify_semantic_artifact(root: Path, protocol: dict[str, Any]) -> str:
    # The capture protocol does not invent an archive exporter.  If root adds
    # an explicit acceptance artifact, its path is declared in protocol and
    # is checked here; otherwise the report gates remain the retained semantic
    # evidence and output wrapper equality stays intentionally out of scope.
    name = protocol.get("semantic_equivalence_artifact")
    if name is None:
        return "report gates retained; no external output wrapper equality required by protocol"
    name = text(name, "protocol.semantic_equivalence_artifact")
    artifact = obj(load(root, name), name)
    if artifact.get("status") != "pass":
        fail(f"{name}: semantic acceptance did not pass")
    for key in ("before", "after"):
        if key in artifact and not isinstance(artifact[key], (dict, list, str)):
            fail(f"{name}.{key}: malformed semantic acceptance row")
    return name


def verify(root: Path = ROOT) -> dict[str, Any]:
    protocol, order = verify_protocol(root)
    before_build = verify_build(root, protocol, "before")
    after_build = verify_build(root, protocol, "after")
    if before_build["revision"] == after_build["revision"]:
        fail("build roles: before and after source revisions unexpectedly match")
    if before_build["binary_sha256"] == after_build["binary_sha256"]:
        fail("build roles: before and after binary identities unexpectedly match")
    harness_files = verify_harness_source_identity(before_build, after_build)
    before, before_samples = verify_role_captures(root, "before", before_build, protocol, order)
    after, after_samples = verify_role_captures(root, "after", after_build, protocol, order)
    verify_cross_role(before, after, order)
    semantic_artifact = verify_semantic_artifact(root, protocol)
    compare_script = safe_path(root, "compare.py")
    run_checked([sys.executable, "-B", str(compare_script), "--check"], "comparison replay")
    check_count = verify_checks(root)
    planned_pass = verify_planned_checks(root)
    compressed = verify_compression(root)
    inventory = verify_inventory(root)
    return {
        "status": "pass",
        "change": CHANGE,
        "roles": 2,
        "processes_per_role": PROCESSES,
        "samples_per_role": before_samples,
        "reports": len(before) + len(after),
        "unchanged_harness_files": harness_files,
        "command_receipts": check_count,
        "planned_required_pass": planned_pass,
        "compressed_logs": compressed,
        "inventory_files": inventory,
        "semantic_acceptance": semantic_artifact,
        "performance_claim": None,
    }


def portable_mutations(root: Path, baseline: dict[str, Any]) -> dict[str, Any]:
    """Exercise report numeric/output and bound-driver custody in a copy."""
    with tempfile.TemporaryDirectory(prefix="litchi-0431-portable-") as temporary:
        exported = Path(temporary) / "bundle"
        shutil.copytree(root, exported)
        # A portable export deliberately has no role executable.  A complete
        # replay must still pass from the copied reports and source manifests.
        # The build records retain the original absolute capture paths as
        # provenance.  They are metadata only; this replay never opens them.
        replay = verify(exported)
        if replay != baseline:
            fail("portable replay result differs from source bundle")

        target_name = "after/plain-bytes-r1.json"
        receipt_name = "after/plain-bytes-r1-receipt.json"
        target_path = safe_path(exported, target_name)
        receipt_path = safe_path(exported, receipt_name)
        original_report = target_path.read_bytes()
        original_receipt = receipt_path.read_bytes()
        # Remove inventory so the mutation reaches the semantic comparison
        # guard instead of stopping at the old sealed digest list.
        inventory_path = exported / "SHA256SUMS"
        inventory_raw = inventory_path.read_bytes() if inventory_path.is_file() else None
        if inventory_path.is_file():
            inventory_path.unlink()
        report = obj(load_bytes(original_report, target_name), target_name)
        row = obj(report["samples_raw"][0], target_name + ".samples_raw[0]")
        timings = obj(row["timings"], target_name + ".timings")
        timings["plan_ns"] += 1
        timings["api_sum_ns"] += 1
        target_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        receipt = obj(load_bytes(original_receipt, receipt_name), receipt_name)
        artifacts = obj(receipt["artifacts"], receipt_name + ".artifacts")
        artifacts[target_name] = {"sha256": digest(target_path.read_bytes()), "bytes": target_path.stat().st_size}
        receipt_path.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
        rejected = subprocess.run([sys.executable, "-B", str(exported / "verify.py")], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if rejected.returncode == 0:
            fail("portable numeric mutation was accepted")
        target_path.write_bytes(original_report)
        receipt_path.write_bytes(original_receipt)

        report = obj(load_bytes(original_report, target_name), target_name)
        original_output = report["expected_output_sha256"]
        replacement = "0" * 64 if original_output != "0" * 64 else "1" * 64
        report["expected_output_sha256"] = replacement
        for raw_row in report["samples_raw"]:
            obj(raw_row, target_name + ".sample")["output_sha256"] = replacement
        target_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        receipt = obj(load_bytes(original_receipt, receipt_name), receipt_name)
        artifacts = obj(receipt["artifacts"], receipt_name + ".artifacts")
        artifacts[target_name] = {"sha256": digest(target_path.read_bytes()), "bytes": target_path.stat().st_size}
        receipt_path.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
        rejected = subprocess.run([sys.executable, "-B", str(exported / "verify.py")], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if rejected.returncode == 0:
            fail("portable output mutation was accepted")
        target_path.write_bytes(original_report)
        receipt_path.write_bytes(original_receipt)
        if inventory_raw is not None:
            inventory_path.write_bytes(inventory_raw)

        verifier = exported / "verify-report.py"
        original_verifier = verifier.read_bytes()
        verifier.write_bytes(original_verifier + b"\n# portable mutation\n")
        rejected = subprocess.run([sys.executable, "-B", str(exported / "verify.py")], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if rejected.returncode == 0:
            fail("portable bound-verifier mutation was accepted")
    return {
        "portable_export": "pass; no executable dependency",
        "portable_numeric_mutation": "rejected",
        "portable_output_mutation": "rejected",
        "portable_bound_verifier_mutation": "rejected",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--portable-check", action="store_true", help="replay and exercise a copied portable bundle")
    args = parser.parse_args(argv)
    try:
        result = verify(ROOT)
        if args.portable_check:
            result.update(portable_mutations(ROOT, result))
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"INVALID: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
