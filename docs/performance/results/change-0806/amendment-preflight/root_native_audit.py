"""Independent custody and numerical audit for the 0806 amendment preflight.

This reader intentionally has no imports from the other amendment readers and
does not execute Cargo, the probe, a profiler, or any other child process.  It
reconstructs the semantic contract and the protected decision from the frozen
JSON reports and then compares every reconstructed row with the primary
analysis.  A result file is written only after every guard passes.
"""

from __future__ import annotations

import json
import math
import os
import random
import re
import statistics
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
PACKET = P.parent
PRIOR = PACKET.parent / "change-0805"
TARGET = Path("/home/zhuhe/code/litchi-target-0806/amendment-preflight")
OUT = P / "root-native-audit.json"

SOURCE_NAMES = {
    "litchi-ole-common-xml_attributes.rs":
        "crates/litchi-ole-common/src/xml_attributes.rs",
    "litchi-opc-xml_attributes.rs": "crates/litchi-opc/src/xml_attributes.rs",
    "litchi-opc-xml_attributes-tests.rs":
        "crates/litchi-opc/src/xml_attributes/tests.rs",
    "litchi-sign-xml_attributes.rs": "crates/litchi-sign/src/xml_attributes.rs",
    "litchi-xldm-xml_attributes.rs": "crates/litchi-xldm/src/xml_attributes.rs",
    "xml-minifier-xml_attributes.rs": "crates/xml-minifier/src/xml_attributes.rs",
}
HELPER_NAMES = tuple(name for name in SOURCE_NAMES if not name.endswith("-tests.rs"))
MODES = ("construct", "consume")
# The probe's source order is [0, 1, 2, 3, 4, 5, 32, 33].  Keep the explicit
# value separate from the tuple above so an accidental edit cannot silently
# make the semantic and capture contracts agree with the wrong order.
CLONE_ADVANCES = (0, 1, 2, 3, 4, 5, 32, 33)
ORDERS = (
    ("before", "after"),
    ("after", "before"),
    ("before", "after"),
    ("after", "before"),
    ("after", "before"),
    ("before", "after"),
)
CPU = 12
BLOCKS = 6
SAMPLES = 30
WARMUP = 3
ITERATIONS = 4096
REPORTS = BLOCKS * 39 * len(MODES) * 2
SAMPLE_COUNT = REPORTS * SAMPLES
BOOTSTRAP_SEED = 806082
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_ENDPOINTS = (250, 9749)
RATIO_THRESHOLD = 1.05
CI_LOW_THRESHOLD = 1.0
BENEFIT_THRESHOLD = 0.97
BENEFIT_CI_HIGH = 1.0

MASK = (1 << 64) - 1
FNV_OFFSET = 0xCBF29CE484222325
FNV_PRIME = 0x00000100000001B3
STEP = 0x9E3779B97F4A7C15
CONSTRUCT_SEED = 0x6A09E667F3BCC909
VALUE_SEED = 0xBB67AE8584CAA73B
VALUE_MULTIPLIER = 0xD6E8FEB86659FD93
AUDIT_SCHEMA = "litchi.performance.0806.amendment-root-native-audit.v1"


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise AssertionError(f"cannot read JSON {path}: {error}") from error


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha(path: Path) -> str:
    import hashlib

    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def is_regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label}: not a regular file")


def descriptor(path: Path, label: str, relative_to: Path | None = None) -> dict[str, Any]:
    is_regular(path, label)
    result: dict[str, Any] = {
        "path": str(path.relative_to(relative_to)) if relative_to is not None else str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }
    return result


def validate_descriptor(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    require(
        isinstance(value.get("path"), str)
        and isinstance(value.get("bytes"), int)
        and not isinstance(value.get("bytes"), bool)
        and value["bytes"] >= 0
        and isinstance(value.get("sha256"), str)
        and re.fullmatch(r"[0-9a-f]{64}", value["sha256"]) is not None,
        f"{label}: malformed descriptor",
    )


def packet_path(raw: Any, label: str, expected: str | None = None) -> Path:
    require(isinstance(raw, str), f"{label}: path is not a string")
    require(not os.path.isabs(raw), f"{label}: packet path must be relative")
    parts = Path(raw).parts
    require(".." not in parts and "." not in parts, f"{label}: unsafe packet path")
    path = (P / raw).resolve()
    require(path.is_relative_to(P.resolve()), f"{label}: packet path escaped packet")
    if expected is not None:
        require(raw == expected, f"{label}: expected {expected!r}, got {raw!r}")
    is_regular(path, label)
    return path


def packet_descriptor(value: Any, expected: str, label: str) -> Path:
    validate_descriptor(value, label)
    path = packet_path(value["path"], label, expected)
    require(value == descriptor(path, label, P), f"{label}: descriptor changed")
    return path


def target_descriptor(value: Any, expected: Path, label: str, cleanup: dict[str, Any] | None) -> None:
    validate_descriptor(value, label)
    require(value["path"] == str(expected), f"{label}: binary path changed")
    if expected.is_file() and not expected.is_symlink():
        require(value == descriptor(expected, label), f"{label}: binary changed")
        return
    require(cleanup is not None, f"{label}: binary disappeared without cleanup witness")
    require(value in cleanup.get("removed_binaries", []), f"{label}: missing cleanup identity")


def u64(value: int) -> int:
    return value & MASK


def rotate_left(value: int, amount: int) -> int:
    return u64((value << amount) | (value >> (64 - amount)))


def fnv(data: bytes) -> int:
    value = FNV_OFFSET
    for byte in data:
        value = u64((value ^ byte) * FNV_PRIME)
    return value


def hash_byte(value: int, byte: int) -> int:
    return u64((value ^ byte) * FNV_PRIME)


def hash_u64(value: int, item: int) -> int:
    for byte in item.to_bytes(8, "little"):
        value = hash_byte(value, byte)
    return value


def error_marker(error: dict[str, Any]) -> int:
    kind = {"ExpectedEq": 1, "ExpectedValue": 2, "UnquotedValue": 3,
            "ExpectedQuote": 4, "Duplicated": 5}[error["kind"]]
    first = error["position"]
    second = error.get("first_position", error.get("quote", 0))
    return u64(kind * STEP) ^ rotate_left(first, 17) ^ rotate_left(second, 31)


def sequence_hash(items: list[dict[str, Any]]) -> int:
    value = FNV_OFFSET
    for item in items:
        if item["kind"] == "Attribute":
            value = hash_byte(value, 1)
            value = hash_u64(value, item["key_bytes"])
            value = hash_u64(value, item["key_hash"])
            value = hash_u64(value, item["value_bytes"])
            value = hash_u64(value, item["value_hash"])
        else:
            value = hash_byte(value, 2)
            value = hash_u64(value, error_marker(item["error"]))
    return value


def expected_case(case: dict[str, Any]) -> dict[str, Any]:
    require(isinstance(case, dict), "case entry is not an object")
    case_id = case.get("id")
    category = case.get("category")
    attribute_count = case.get("attribute_count")
    require(isinstance(case_id, str) and isinstance(category, str), "case identity malformed")

    if category == "distinct":
        require(case_id.startswith("distinct-") and isinstance(attribute_count, int),
                f"{case_id}: distinct identity malformed")
        count = attribute_count
        expected_id = f"distinct-{count}"
        tail = ""
        error = None
    elif category in {
        "duplicate-valid", "duplicate-long-quoted", "duplicate-long-unterminated",
        "duplicate-unquoted",
    }:
        prefix_name, separator, suffix = case_id.rpartition("-after-")
        require(separator and prefix_name == category and suffix.isdigit(),
                f"{case_id}: duplicate identity malformed")
        count = int(suffix)
        expected_id = f"{category}-after-{count}"
        if category == "duplicate-valid":
            tail = 'n0="again"'
        elif category == "duplicate-long-quoted":
            tail = 'n0="' + "x" * 4096 + '"'
        elif category == "duplicate-long-unterminated":
            tail = 'n0="' + "x" * 4096
        else:
            tail = "n0=" + "x" * 4096
        error = {"kind": "Duplicated", "position": len(_distinct_input(count)) + 1,
                 "first_position": 2}
        require(attribute_count == count + 1, f"{case_id}: attribute count changed")
    elif category in {"syntax-flag", "syntax-unique-tail", "syntax-equals-value"}:
        prefix_name, separator, suffix = case_id.rpartition("-after-")
        require(separator and prefix_name == category and suffix.isdigit(),
                f"{case_id}: syntax identity malformed")
        count = int(suffix)
        expected_id = f"{category}-after-{count}"
        tail = {
            "syntax-flag": "flag",
            "syntax-unique-tail": "tail=x",
            "syntax-equals-value": '="1"',
        }[category]
        if category == "syntax-unique-tail":
            error = {"kind": "UnquotedValue", "position": len(_distinct_input(count)) + 6}
        else:
            error = {"kind": "ExpectedEq", "position": len(_distinct_input(count) + " " + tail)}
        require(attribute_count == count, f"{case_id}: attribute count changed")
    else:
        raise AssertionError(f"{case_id}: unknown case category")

    require(case_id == expected_id, f"{case_id}: case identity is not canonical")
    source = _distinct_input(count) if not tail else _distinct_input(count) + " " + tail
    source_bytes = source.encode("utf-8")
    source_identity = {
        "bytes": len(source_bytes),
        "encoding": "utf8" if len(source_bytes) <= 512 else "hex",
        "value": source if len(source_bytes) <= 512 else source_bytes.hex(),
    }
    require(case.get("source") == source_identity, f"{case_id}: source identity changed")

    items: list[dict[str, Any]] = []
    for index in range(count):
        key = f"n{index}".encode()
        value = str(index).encode()
        items.append({
            "kind": "Attribute",
            "key_bytes": len(key),
            "key_hash": fnv(key),
            "value_bytes": len(value),
            "value_hash": fnv(value),
        })
    if error is not None:
        items.append({"kind": "Error", "error": error})
    trace = {
        "accepted": count,
        "first_error": error,
        "repeated_none_calls": 2,
        "sequence_hash": sequence_hash(items),
        "items": items,
    }
    require(case.get("expected_baseline") == trace, f"{case_id}: semantic fixture changed")
    return {
        "id": case_id,
        "category": category,
        "attribute_count": attribute_count,
        "source": source_identity,
        "trace": trace,
        "items": items,
        "error": error,
    }


def _distinct_input(count: int) -> str:
    return "e" if count == 0 else "e " + " ".join(
        f'n{index}="{index}"' for index in range(count)
    )


def expected_clone_hash(items: list[dict[str, Any]], advance: int) -> int:
    if any(item["kind"] == "Error" for item in items):
        error_index = next(index for index, item in enumerate(items)
                           if item["kind"] == "Error")
        suffix = items[advance:] if advance <= error_index else []
    else:
        suffix = items[advance:]
    return sequence_hash(suffix)


def expected_result(trace: dict[str, Any], mode: str) -> dict[str, int]:
    if mode == "construct":
        checksum = CONSTRUCT_SEED
        for index in range(ITERATIONS):
            checksum = u64(checksum + u64((index + 1) * STEP))
        return {"checksum": checksum, "accepted": 0, "error_marker": 0}

    one_checksum = VALUE_SEED
    one_accepted = 0
    one_error = 0
    for item in trace["items"]:
        if item["kind"] == "Attribute":
            one_accepted += 1
            one_checksum = rotate_left(one_checksum, 7) ^ (
                rotate_left(VALUE_SEED, 5)
                ^ u64(item["key_bytes"] * STEP)
                ^ u64(item["value_bytes"] * VALUE_MULTIPLIER)
            )
        else:
            one_error = error_marker(item["error"])

    checksum = CONSTRUCT_SEED
    accepted = 0
    error = 0
    for index in range(ITERATIONS):
        checksum = u64(checksum + (one_checksum ^ u64((index + 1) * STEP)))
        accepted = u64(accepted + one_accepted)
        error = u64(error + one_error)
    return {"checksum": checksum, "accepted": accepted, "error_marker": error}


def validate_contract() -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    plan = read_json(P / "plan.json")
    require(plan.get("schema") == "litchi.performance.0806.amendment-preflight.v1",
            "amendment plan schema changed")
    require(plan.get("cpu") == CPU and plan.get("legs") == ["before", "after"]
            and plan.get("modes") == list(MODES) and plan.get("case_count") == 39,
            "amendment schedule changed")
    native = plan.get("native")
    require(native == {
        "blocks": BLOCKS,
        "orders": [list(order) for order in ORDERS],
        "samples": SAMPLES,
        "warmup": WARMUP,
        "iterations": ITERATIONS,
    }, "native contract changed")
    analysis = plan.get("analysis")
    require(analysis == {
        "bootstrap_seed": BOOTSTRAP_SEED,
        "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        "bootstrap_statistic": "median of six paired process-p50 ratios",
        "process_p50": "nearest rank ceil(n/2)-1",
        "zero_based_endpoints": list(BOOTSTRAP_ENDPOINTS),
        "diagnostic_ratio_above": RATIO_THRESHOLD,
        "diagnostic_ci_low_above": CI_LOW_THRESHOLD,
    }, "bootstrap contract changed")
    scope = plan.get("scope")
    require(isinstance(scope, dict)
            and scope.get("claim") == "protected native micro-input timing only"
            and scope.get("public_workflow_speedup") is False
            and scope.get("resource_claim") is False
            and scope.get("production_adoption") is False
            and scope.get("profiles") is False
            and scope.get("callgrind") is False
            and scope.get("allocator_measurement") is False
            and scope.get("native_process_elapsed_is_diagnostic") is True,
            "scope contract changed")
    lineage = plan.get("lineage")
    require(isinstance(lineage, dict)
            and lineage.get("original_production_before")
            == "../candidate/before"
            and lineage.get("amended_candidate_after")
            == "../candidate-quality-amendment/after"
            and lineage.get("comparison") == "source/before versus source/after",
            "amendment source lineage changed")
    policy = plan.get("policy")
    require(isinstance(policy, dict)
            and policy.get("benefit_mode") == "consume"
            and policy.get("benefit_cases") == ["distinct-1", "distinct-2"]
            and policy.get("benefit_ratio_at_most") == BENEFIT_THRESHOLD
            and policy.get("benefit_ci_high_below") == BENEFIT_CI_HIGH
            and policy.get("protected_ratio_above") == RATIO_THRESHOLD
            and policy.get("failure_action")
            == "retain both archives and reject amendment; do not run public workflow adoption captures"
            and policy.get("success_action")
            == "amendment eligible only for root review; this packet never adopts production",
            "policy contract changed")
    protected = policy.get("protected_consume_cases")
    require(protected == [
        "distinct-0", "distinct-1", "distinct-2",
        "duplicate-valid-after-1", "duplicate-valid-after-2",
        "duplicate-long-quoted-after-1", "duplicate-long-quoted-after-2",
        "duplicate-long-quoted-after-33", "duplicate-long-unterminated-after-1",
        "duplicate-long-unterminated-after-2", "duplicate-long-unterminated-after-33",
        "duplicate-unquoted-after-1", "syntax-flag-after-0", "syntax-flag-after-2",
        "syntax-unique-tail-after-0", "syntax-unique-tail-after-2",
        "syntax-equals-value-after-0", "syntax-equals-value-after-2",
    ], "protected policy cases changed")

    manifest = read_json(P / "manifest.json")
    require(manifest.get("schema") == "litchi.performance.0806.amendment-preflight-manifest.v1"
            and manifest.get("packet") == "change-0806/amendment-preflight",
            "amendment manifest changed")
    source_spec = manifest.get("source")
    require(isinstance(source_spec, dict)
            and source_spec.get("before") == "source/before"
            and source_spec.get("after") == "source/after"
            and source_spec.get("before_lineage")
            == "../candidate/before"
            and source_spec.get("after_lineage")
            == "../candidate-quality-amendment/after"
            and isinstance(source_spec.get("handoff_before_is_not_used"), str),
            "source lineage contract changed")
    probe = manifest.get("probe")
    require(isinstance(probe, dict)
            and probe.get("schema") == "litchi.attribute-boundary-probe.v1"
            and probe.get("tool") == "attribute-boundary-probe-0805"
            and probe.get("clone_advances") == list(CLONE_ADVANCES),
            "probe contract changed")
    execution = manifest.get("execution")
    require(execution == {
        "target": str(TARGET),
        "cpu": CPU,
        "native_blocks": BLOCKS,
        "native_samples": SAMPLES,
        "native_warmup": WARMUP,
        "native_iterations": ITERATIONS,
        "bootstrap_seed": BOOTSTRAP_SEED,
        "profiles": False,
        "callgrind": False,
        "historical_timing_pooling": False,
    }, "execution contract changed")

    cases = read_json(P / "cases.json")
    fixtures = read_json(P / "fixtures.json")
    old_cases = read_json(PRIOR / "cases.json")
    old_fixtures = read_json(PRIOR / "fixtures.json")
    require(cases == old_cases and fixtures == old_fixtures,
            "sealed 0805 fixture archive changed")
    require(isinstance(cases, list) and len(cases) == 39
            and isinstance(fixtures, list) and len(fixtures) == 39,
            "case or fixture count changed")
    expected_cases = [expected_case(case) for case in cases]
    require([case["id"] for case in expected_cases]
            == [case["id"] for case in cases], "case order changed")
    fixture_by_id = {fixture.get("id"): fixture for fixture in fixtures}
    require(set(fixture_by_id) == {case["id"] for case in expected_cases},
            "fixture identity changed")
    for case in cases:
        fixture = fixture_by_id[case["id"]]
        expected_source = case["source"]
        encoded = (expected_source["value"] if expected_source["encoding"] == "utf8"
                   else bytes.fromhex(expected_source["value"]).decode("utf-8"))
        require(fixture.get("category") == case["category"]
                and fixture.get("attribute_count") == case["attribute_count"]
                and fixture.get("input") == encoded
                and fixture.get("bytes") == expected_source["bytes"]
                and isinstance(fixture.get("sha256"), str)
                and re.fullmatch(r"[0-9a-f]{64}", fixture["sha256"]) is not None,
                f"{case['id']}: fixture contract changed")
        require(sha256_bytes(encoded.encode()) == fixture["sha256"],
                f"{case['id']}: fixture hash changed")
    return plan, cases, {"manifest": manifest, "expected_cases": expected_cases}


def sha256_bytes(value: bytes) -> str:
    import hashlib

    return hashlib.sha256(value).hexdigest()


def archive_manifest(leg: str) -> dict[str, Any]:
    directory = P / "source" / leg
    require(directory.is_dir() and not directory.is_symlink(), f"missing source archive {leg}")
    expected_names = set(SOURCE_NAMES)
    actual = {path.name for path in directory.iterdir() if path.is_file() and not path.is_symlink()}
    require(actual == expected_names, f"{leg}: source inventory changed")
    return {
        name: descriptor(directory / name, f"source/{leg}/{name}", P)
        for name in sorted(SOURCE_NAMES)
    }


def validate_sources() -> dict[str, Any]:
    amendment = read_json(PACKET / "candidate-quality-amendment" / "manifest.json")
    require(amendment.get("schema") == "litchi.performance.0806.quality-amendment.v1",
            "quality amendment manifest changed")
    before_root = PACKET / "candidate" / "before"
    after_root = PACKET / "candidate-quality-amendment" / "after"
    handoff_before_root = PACKET / "candidate-quality-amendment" / "before"
    candidate_after = PACKET / "candidate" / "after"
    source_hashes: dict[str, dict[str, str]] = {}
    for leg, expected_root in (("before", before_root), ("after", after_root)):
        source_hashes[leg] = {}
        for name in SOURCE_NAMES:
            archive = P / "source" / leg / name
            expected = expected_root / name
            if leg == "after" and name.endswith("-tests.rs"):
                # The five-file quality-amendment handoff omits the shared
                # OPC tests; its byte-identical witness is candidate/after.
                expected = candidate_after / name
            is_regular(archive, f"source/{leg}/{name}")
            is_regular(expected, f"lineage/{leg}/{name}")
            require(sha(archive) == sha(expected), f"{leg}/{name}: lineage changed")
            source_hashes[leg][name] = sha(archive)
    test_name = "litchi-opc-xml_attributes-tests.rs"
    require(source_hashes["before"][test_name]
            == sha(PACKET / "candidate" / "before" / test_name),
            "before shared OPC test archive drifted")
    require(source_hashes["after"][test_name] == sha(candidate_after / test_name),
            "after shared OPC test archive drifted")

    old = (
        "    #[allow(clippy::disallowed_methods)]\n"
        "    #[inline]\n"
        "    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
        "        let mut attributes = tag.attributes();\n"
        "        attributes.with_checks(false);"
    )
    new = (
        "    #[inline]\n"
        "    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
        "        let attributes = tag.unchecked_attributes();"
    )
    for name in HELPER_NAMES:
        # The native lane's before side is the original workflow baseline.
        # The mechanical quality amendment is independently checked against
        # the handoff's own before copy, which is the already-applied 0806
        # candidate and is intentionally excluded from the timed comparison.
        before = (handoff_before_root / name).read_text(encoding="utf-8")
        after = (after_root / name).read_text(encoding="utf-8")
        require(before.count(old) == 1 and after == before.replace(old, new, 1),
                f"{name}: amendment is not the frozen constructor rewrite")
        require(after.count("let attributes = tag.unchecked_attributes();") == 1,
                f"{name}: helper call missing")
        require(after.count("attributes.with_checks(false);") == 1,
                f"{name}: helper implementation changed")

    # Bind the amendment manifest's five after descriptors and shared-test
    # witness independently of the primary reader.
    files = amendment.get("files")
    require(isinstance(files, dict) and set(files) == set(HELPER_NAMES),
            "quality amendment file inventory changed")
    for name in HELPER_NAMES:
        row = files[name]
        require(row.get("production_path") == SOURCE_NAMES[name],
                f"{name}: production path changed")
        for side, root in (("before", handoff_before_root), ("after", after_root)):
            value = row.get(side)
            require(isinstance(value, dict)
                    and value.get("path") == str(root / name)
                    and value.get("bytes") == (root / name).stat().st_size
                    and value.get("sha256") == sha(root / name),
                    f"{name}: amendment {side} witness changed")
    shared = amendment.get("shared_files", {}).get("litchi-opc-xml_attributes-tests.rs", {})
    require(shared.get("original_candidate_after", {}).get("sha256")
            == source_hashes["after"][test_name]
            and shared.get("amendment_action") == "byte-identical; omitted from this five-helper amendment",
            "amendment shared-test witness changed")
    return {"archives": {"before": archive_manifest("before"),
                          "after": archive_manifest("after")},
            "hashes": source_hashes}


def expected_probe_manifest() -> dict[str, str]:
    directory = P / "probe-src"
    require(directory.is_dir() and not directory.is_symlink(), "probe source missing")
    expected = {
        "Cargo.lock",
        "Cargo.toml.template",
        "src/baseline.rs",
        "src/candidate.rs",
        "src/main.rs",
    }
    materialized = directory / "Cargo.toml"
    for path in directory.rglob("*"):
        require(not path.is_symlink(), f"probe/{path.relative_to(directory)}: symlink is forbidden")
    actual = {
        str(path.relative_to(directory))
        for path in directory.rglob("*")
        if path.is_file() and not path.is_symlink() and path.name != "Cargo.toml"
    }
    require(actual == expected, f"probe inventory changed: {sorted(actual)}")
    result = {f"probe-src/{name}": sha(directory / name) for name in sorted(expected)}
    for name in expected:
        is_regular(directory / name, f"probe/{name}")
    require(result["probe-src/Cargo.lock"] == sha(PRIOR / "probe-src" / "Cargo.lock"),
            "probe lock differs from sealed 0805")
    require(result["probe-src/Cargo.toml.template"] == sha(PRIOR / "probe-src" / "Cargo.toml.template"),
            "probe manifest template differs from sealed 0805")
    require(result["probe-src/src/main.rs"] == sha(PRIOR / "probe-src" / "src/main.rs"),
            "probe driver differs from sealed 0805")
    source_before = P / "source" / "before" / "litchi-opc-xml_attributes.rs"
    source_after = P / "source" / "after" / "litchi-opc-xml_attributes.rs"
    require(result["probe-src/src/baseline.rs"] == sha(source_before), "baseline probe source drifted")
    require(result["probe-src/src/candidate.rs"] == sha(source_after), "candidate probe source drifted")
    template = (directory / "Cargo.toml.template").read_text(encoding="utf-8")
    if materialized.exists():
        require((directory / "Cargo.toml").read_text(encoding="utf-8")
                == template.replace("@SRC@", str(P.parents[4])),
                "materialized probe manifest changed")
    return result


def validate_cleanup() -> dict[str, Any] | None:
    path = P / "cleanup.json"
    if not path.exists():
        return None
    cleanup = read_json(path)
    require(cleanup.get("schema") == "litchi.performance.0806.amendment-cleanup.v1"
            and cleanup.get("target") == str(TARGET)
            and cleanup.get("target_removed") is True
            and not TARGET.exists(), "amendment cleanup witness changed")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 2, "cleanup binary count changed")
    return cleanup


def validate_builds(archives: dict[str, Any], probe: dict[str, str],
                    cleanup: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    builds: dict[str, dict[str, Any]] = {}
    source_values: dict[str, dict[str, Any]] = {}
    for leg in ("before", "after"):
        root = P / f"build-{leg}"
        build = read_json(root / "build.json")
        require(build.get("schema") == "litchi.performance.0806.amendment-build.v1"
                and build.get("leg") == leg
                and build.get("archive") == archives[leg]
                and build.get("probe") == probe
                and build.get("profiles") is False
                and build.get("callgrind") is False,
                f"{leg}: build custody changed")
        source = build.get("source")
        require(isinstance(source, dict) and isinstance(source.get("revision"), str)
                and isinstance(source.get("files"), dict), f"{leg}: source census missing")
        source_values[leg] = source
        frozen = read_json(root / "frozen-inputs.json")
        require(frozen.get("schema") == "litchi.performance.0806.amendment-build-inputs.v1"
                and frozen.get("leg") == leg
                and frozen.get("archives") == {"before": archives["before"], "after": archives["after"]}
                and frozen.get("probe") == probe
                and frozen.get("workspace_source") == source,
                f"{leg}: frozen build inputs changed")
        require(frozen.get("plan") == sha(P / "plan.json")
                and frozen.get("build_driver") == sha(P / "build.py")
                and frozen.get("quality_driver") == sha(P / "quality.py")
                and frozen.get("capture_driver") == sha(P / "capture.py")
                and frozen.get("custody_driver") == sha(P / "custody.py"),
                f"{leg}: frozen driver witness changed")
        command = build.get("command")
        expected_command = [
            "cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", str(P / "probe-src" / "Cargo.toml"),
            "--bin", "attribute-boundary-probe",
        ]
        require(isinstance(command, dict) and command.get("command") == expected_command
                and command.get("exit_code") == 0
                and isinstance(command.get("started"), (int, float))
                and isinstance(command.get("ended"), (int, float))
                and command["started"] <= command["ended"],
                f"{leg}: build command receipt changed")
        packet_descriptor(command.get("log"), f"build-{leg}/native.log", f"{leg}: build log")
        require(read_json(root / "commands.json") == [command],
                f"{leg}: command receipt disagrees")
        binary = build.get("binary")
        target_descriptor(binary, TARGET / f"{leg}-native", f"{leg}: native binary", cleanup)
        lock = build.get("lock")
        validate_descriptor(lock, f"{leg}: lock")
        require(lock.get("path") == str(P / "probe-src" / "Cargo.lock")
                and lock == descriptor(P / "probe-src" / "Cargo.lock", f"{leg}: lock"),
                f"{leg}: lock witness changed")
        builds[leg] = build

    before_source = source_values["before"]
    after_source = source_values["after"]
    require(before_source == after_source,
            "build workspace source changed between legs")
    candidate_manifest = read_json(PACKET / "candidate" / "manifest.json")
    require(before_source["revision"] == candidate_manifest.get("base_commit"),
            "build source revision is not the frozen production revision")
    application = read_json(PACKET / "application.json")
    require(before_source == application.get("source"),
            "build source census does not match production-before witness")
    require(command_time(builds["before"]) <= command_time(builds["after"]),
            "after build did not follow before build")
    return builds


def command_time(build: dict[str, Any]) -> float:
    command = build["command"]
    return float(command["ended"])


def expected_receipt_command(leg: str, case_id: str, mode: str, stem: str) -> list[str]:
    report = P / "native" / f"{stem}.json"
    rss = P / "native" / f"{stem}.rss"
    return [
        "/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c", str(CPU),
        str(TARGET / f"{leg}-native"), "--leg", leg, "--case", case_id,
        "--mode", mode, "--samples", str(SAMPLES), "--warmup", str(WARMUP),
        "--iterations", str(ITERATIONS), "--output", str(report),
    ]


def validate_report(report: dict[str, Any], expected: dict[str, Any], leg: str,
                    mode: str, case_id: str, receipt_index: int) -> None:
    label = f"native report {receipt_index}"
    require(report.get("schema") == "litchi.attribute-boundary-probe.v1"
            and report.get("tool") == "attribute-boundary-probe-0805"
            and report.get("binary") == f"{leg}-native"
            and report.get("leg") == leg
            and report.get("case") == case_id
            and report.get("category") == expected["category"]
            and report.get("attribute_count") == expected["attribute_count"]
            and report.get("mode") == mode
            and report.get("iterations") == ITERATIONS
            and report.get("warmup") == WARMUP
            and report.get("samples_requested") == SAMPLES,
            f"{label}: report metadata changed")
    timing_scope = {
        "construct": "selected named construction owner call, including its common call dispatch",
        "consume": "selected named consumption owner call, including its common call dispatch",
    }[mode]
    require(report.get("timing_scope") == timing_scope, f"{label}: timing scope changed")
    require(report.get("source") == expected["source"], f"{label}: source identity changed")
    oracle = report.get("semantic_oracle")
    require(isinstance(oracle, dict)
            and oracle.get("quick_xml") == expected["trace"]
            and oracle.get("baseline") == expected["trace"]
            and oracle.get("candidate") == expected["trace"]
            and oracle.get("baseline_matches_quick_xml") is True
            and oracle.get("candidate_matches_quick_xml") is True
            and oracle.get("all_checks_passed") is True,
            f"{label}: semantic oracle changed")
    clones = oracle.get("clone_checks")
    require(isinstance(clones, list) and len(clones) == len(CLONE_ADVANCES),
            f"{label}: clone count changed")
    for clone, advance in zip(clones, CLONE_ADVANCES):
        expected_hash = expected_clone_hash(expected["items"], advance)
        require(clone == {
            "advance": advance,
            "quick_xml_sequence_hash": expected_hash,
            "baseline_sequence_hash": expected_hash,
            "candidate_sequence_hash": expected_hash,
            "baseline_matches_quick_xml": True,
            "candidate_matches_quick_xml": True,
            "terminal_behavior_matches": True,
        }, f"{label}: clone oracle changed at {advance}")
    require(report.get("iterator_sizes") == {
        "baseline_checked_attributes": 120,
        "candidate_checked_attributes": 128,
    }, f"{label}: iterator layout changed")
    result = expected_result(expected["trace"], mode)
    require(report.get("expected_result") == result, f"{label}: expected result changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES,
            f"{label}: sample count changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict)
                and sample.get("index") == index
                and isinstance(sample.get("elapsed_ns"), int)
                and not isinstance(sample.get("elapsed_ns"), bool)
                and sample["elapsed_ns"] > 0
                and sample.get("checksum") == result["checksum"]
                and sample.get("accepted") == result["accepted"]
                and sample.get("error_marker") == result["error_marker"],
                f"{label}: sample {index} changed")


def nearest_rank_p50(samples: list[int]) -> int:
    require(len(samples) == SAMPLES, "invalid p50 sample count")
    ordered = sorted(samples)
    rank_index = (len(ordered) + 1) // 2 - 1
    require(0 <= rank_index < len(ordered), "invalid nearest-rank index")
    return ordered[rank_index]


def bootstrap(values: list[float]) -> tuple[float, float, float]:
    require(len(values) == BLOCKS and all(math.isfinite(value) and value > 0 for value in values),
            "invalid paired ratio vector")
    rng = random.Random(BOOTSTRAP_SEED)
    medians = [statistics.median(values[rng.randrange(len(values))] for _ in values)
               for _ in range(BOOTSTRAP_RESAMPLES)]
    medians.sort()
    return statistics.median(values), medians[BOOTSTRAP_ENDPOINTS[0]], medians[BOOTSTRAP_ENDPOINTS[1]]


def validate_capture(builds: dict[str, dict[str, Any]], archives: dict[str, Any],
                     probe: dict[str, str], expected_cases: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    complete = read_json(P / "native" / "complete.json")
    require(complete.get("schema") == "litchi.performance.0806.amendment-native.v1"
            and complete.get("children") == REPORTS
            and complete.get("expected_children") == REPORTS
            and complete.get("samples") == SAMPLE_COUNT
            and complete.get("profiles") is False
            and complete.get("callgrind") is False
            and complete.get("historical_timing_pooling") is False,
            "native completion contract changed")
    packet_descriptor(complete.get("receipts"), "native/receipts.json", "native receipts")
    packet_descriptor(complete.get("source"), "native/source.json", "native source")
    source = read_json(P / "native" / "source.json")
    require(source.get("schema") == "litchi.performance.0806.amendment-capture-source.v1"
            and source.get("archives") == archives
            and source.get("probe") == probe
            and source.get("workspace") == builds["after"]["source"],
            "native source witness changed")
    receipts = read_json(P / "native" / "receipts.json")
    require(isinstance(receipts, list) and len(receipts) == REPORTS,
            "native receipt count changed")
    case_map = {case["id"]: case for case in expected_cases}
    rows: dict[tuple[str, str, int, str], dict[str, Any]] = {}
    previous_end: float | None = None
    for index, receipt in enumerate(receipts):
        block = index // (39 * len(MODES) * 2)
        local = index % (39 * len(MODES) * 2)
        case_id = expected_cases[local // 4]["id"]
        mode = MODES[(local % 4) // 2]
        leg = ORDERS[block][local % 2]
        stem = f"{block}-{case_id}-{mode}-{leg}"
        label = f"native receipt {index}"
        require(receipt.get("schema") == "litchi.performance.0806.amendment-native-receipt.v1"
                and receipt.get("block") == block
                and receipt.get("case") == case_id
                and receipt.get("mode") == mode
                and receipt.get("leg") == leg
                and receipt.get("exit_code") == 0,
                f"{label}: identity or exit status changed")
        started = receipt.get("started")
        ended = receipt.get("ended")
        require(isinstance(started, (int, float)) and not isinstance(started, bool)
                and isinstance(ended, (int, float)) and not isinstance(ended, bool)
                and math.isfinite(float(started)) and math.isfinite(float(ended))
                and started <= ended
                and (previous_end is None or previous_end <= started),
                f"{label}: chronology changed")
        previous_end = float(ended)
        require(receipt.get("binary") == builds[leg]["binary"],
                f"{label}: binary binding changed")
        packet_descriptor(receipt.get("log"), f"native/{stem}.log", f"{label} log")
        report_path = packet_descriptor(receipt.get("report"), f"native/{stem}.json", f"{label} report")
        rss_path = packet_descriptor(receipt.get("rss"), f"native/{stem}.rss", f"{label} RSS")
        rss_text = rss_path.read_text(encoding="utf-8").strip()
        require(rss_text.isdigit() and int(rss_text) > 0, f"{label}: RSS receipt changed")
        require(receipt.get("command") == expected_receipt_command(leg, case_id, mode, stem),
                f"{label}: command changed")
        report = read_json(report_path)
        expected = case_map[case_id]
        validate_report(report, expected, leg, mode, case_id, index)
        key = (case_id, mode, block, leg)
        require(key not in rows, f"{label}: duplicate identity")
        rows[key] = report
    require(len(rows) == REPORTS, "native report identities are incomplete")
    return [rows[(case["id"], mode, block, leg)]
            for case in expected_cases for mode in MODES for block in range(BLOCKS)
            for leg in ("before", "after")], source


def analyze_rows(expected_cases: list[dict[str, Any]], reports: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key: dict[tuple[str, str, int, str], dict[str, Any]] = {}
    index = 0
    for case in expected_cases:
        for mode in MODES:
            for block in range(BLOCKS):
                for leg in ("before", "after"):
                    by_key[(case["id"], mode, block, leg)] = reports[index]
                    index += 1
    rows: list[dict[str, Any]] = []
    for case in expected_cases:
        for mode in MODES:
            before_p50: list[int] = []
            after_p50: list[int] = []
            ratios: list[float] = []
            for block in range(BLOCKS):
                before = nearest_rank_p50([sample["elapsed_ns"]
                                           for sample in by_key[(case["id"], mode, block, "before")]["samples"]])
                after = nearest_rank_p50([sample["elapsed_ns"]
                                         for sample in by_key[(case["id"], mode, block, "after")]["samples"]])
                require(before > 0, f"{case['id']}/{mode}/{block}: zero before p50")
                before_p50.append(before)
                after_p50.append(after)
                ratios.append(after / before)
            ratio, low, high = bootstrap(ratios)
            rows.append({
                "case": case["id"],
                "mode": mode,
                "process_p50_before": before_p50,
                "process_p50_after": after_p50,
                "paired_ratios": ratios,
                "ratio_median": ratio,
                "bootstrap_ci_low": low,
                "bootstrap_ci_high": high,
                "change_percent_median": (ratio - 1.0) * 100.0,
                # The primary reader records the generic diagnostic flag for
                # every mode and case.  Protected eligibility applies the
                # separate mode/case filter below; keeping those predicates
                # distinct lets this audit compare every row exactly.
                "diagnostic_regression": (
                    ratio > RATIO_THRESHOLD and low > CI_LOW_THRESHOLD
                ),
            })
    return rows


def validate_primary(rows: list[dict[str, Any]], plan: dict[str, Any]) -> bool:
    primary = read_json(P / "analysis.json")
    require(primary.get("schema") == "litchi.performance.0806.amendment-analysis.v1",
            "primary amendment analysis schema changed")
    require(primary.get("policy") == plan["policy"] and primary.get("claims") == plan["scope"],
            "primary amendment analysis contract changed")
    require(primary.get("rows") == rows, "primary analysis does not match independent rows")
    return True


def audit() -> dict[str, Any]:
    plan, cases, contract = validate_contract()
    sources = validate_sources()
    probe = expected_probe_manifest()
    cleanup = validate_cleanup()
    builds = validate_builds(sources["archives"], probe, cleanup)
    reports, _capture_source = validate_capture(builds, sources["archives"], probe,
                                                 contract["expected_cases"])
    rows = analyze_rows(contract["expected_cases"], reports)
    matches_primary = validate_primary(rows, plan)
    protected = [row for row in rows if row["mode"] == "consume"
                 and row["case"] in set(plan["policy"]["protected_consume_cases"])
                 and row["diagnostic_regression"]]
    benefits = {
        case_id: next(row for row in rows
                      if row["case"] == case_id and row["mode"] == "consume")["ratio_median"] <= BENEFIT_THRESHOLD
        and next(row for row in rows
                 if row["case"] == case_id and row["mode"] == "consume")["bootstrap_ci_high"] < BENEFIT_CI_HIGH
        for case_id in plan["policy"]["benefit_cases"]
    }
    advance = bool(all(benefits.values()) and not protected)
    require(matches_primary, "primary analysis comparison did not pass")
    return {
        "schema": AUDIT_SCHEMA,
        "passed": True,
        "native_reports": REPORTS,
        "native_samples": SAMPLE_COUNT,
        "advance_to_workflow_trials": advance,
        "production_adoption": False,
        "matches_primary_analysis": matches_primary,
        "protected_consume_regressions": protected,
        "dominant_class_benefits": benefits,
        "rows": rows,
        "contract": {
            "case_count": len(cases),
            "modes": list(MODES),
            "blocks": BLOCKS,
            "samples": SAMPLES,
            "warmup": WARMUP,
            "iterations": ITERATIONS,
            "cpu": CPU,
            "clone_advances": list(CLONE_ADVANCES),
            "bootstrap_seed": BOOTSTRAP_SEED,
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
            "nearest_rank_p50": "sorted samples[(n+1)//2-1]",
            "profiler_lane": False,
        },
        "source_witness": {
            "archives": sources["archives"],
            "before_mirrors": "candidate/before",
            "after_mirrors": "candidate-quality-amendment/after",
            "probe": probe,
        },
        "build_witness": {
            leg: {
                "build": descriptor(P / f"build-{leg}" / "build.json", f"{leg} build", P),
                "binary": builds[leg]["binary"],
                "source_revision": builds[leg]["source"]["revision"],
            }
            for leg in ("before", "after")
        },
        "capture_witness": {
            "complete": descriptor(P / "native" / "complete.json", "native complete", P),
            "receipts": descriptor(P / "native" / "receipts.json", "native receipts", P),
            "source": descriptor(P / "native" / "source.json", "native source", P),
        },
        "auditor": descriptor(Path(__file__), "root native auditor", P),
    }


def main() -> None:
    check = "--check" in sys.argv[1:]
    result = audit()
    if check:
        require(OUT.is_file(), "root-native-audit.json is missing")
        require(read_json(OUT) == result, "root-native-audit.json does not reproduce")
    else:
        require(not OUT.exists(), "root-native-audit.json already exists")
        OUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print("Independent 0806 amendment native custody, semantic, policy, and numerical audit PASS")


if __name__ == "__main__":
    main()
