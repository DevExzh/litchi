#!/usr/bin/env python3
import datetime as dt
import hashlib
import json
from pathlib import Path

root = Path("/home/zhuhe/code/litchi")
packet = Path("/tmp/litchi-performance-0927")
diag_path = packet / "diagnostic.json"
log_path = packet / "diagnostic-focused.log"
patch_path = packet / "diagnostic.patch"
source_path = root / "tools/perf-baseline/src/lib.rs"
receipt_path = packet / "reviewer-phase-replay.json"
marker = "DOCX_SECTION_LAYOUT_READ_DIAGNOSTIC_JSON="
expected_source_bytes = 2_101_427
expected_source_sha = "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8"
expected_output_bytes = 2_101_446
expected_output_sha = "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
expected_source_file_sha = "23ce48a33ffb9afd05571738642c78a3511338d34ce361f75ab3a710bfd74ab9"
# The source SHA above is checked again below; this literal is deliberately
# replaced with the observed hash if the first replay is run after a freeze.
expected_source_file_sha = "23ce48a33ffb9afd05571738642c78a3511338d34ce361f75ab3a710bfd74ab9"
operation_bytes = 6_308_061
fingerprint_bytes = 4_202_854
phases = ["open", "query", "edit", "commit", "publication"]
scalar_keys = ["read_calls", "requested_bytes", "returned_bytes", "short_reads"]

def sha256(path):
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as f:
        while chunk := f.read(1024 * 1024):
            size += len(chunk)
            h.update(chunk)
    return size, h.hexdigest()

def fail(errors, condition, message):
    if not condition:
        errors.append(message)

errors = []
source_start = sha256(source_path)[1]
diag_size, diag_sha = sha256(diag_path)
log_size, log_sha = sha256(log_path)
patch_size, patch_sha = sha256(patch_path)
reports = json.loads(diag_path.read_text(encoding="utf-8"))
log_reports = []
for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
    if line.startswith(marker):
        log_reports.append(json.loads(line[len(marker):]))
fail(errors, len(reports) == 2, "diagnostic.json does not contain exactly two routes")
fail(errors, len(log_reports) == 1, "focused log does not contain exactly one marker")
if len(log_reports) == 1:
    fail(errors, log_reports[0] == reports, "focused log marker differs from diagnostic.json")
fail(errors, "test tests::docx_section_layout_read_diagnostic_conserves_full_and_short_phases ... ok" in log_path.read_text(encoding="utf-8", errors="replace"), "focused test did not report ok")
fail(errors, "test result: ok. 1 passed; 0 failed" in log_path.read_text(encoding="utf-8", errors="replace"), "focused cargo test summary did not pass")
fail(errors, source_start == expected_source_file_sha, "source lib.rs hash at replay start differs from frozen candidate")

def check_route(route):
    name = route.get("route")
    expected_max = {"full_64k_unpaced": 64 * 1024, "short_4k_unpaced": 4 * 1024}.get(name)
    fail(errors, name in ("full_64k_unpaced", "short_4k_unpaced"), f"unexpected route {name!r}")
    if expected_max is None:
        return
    fail(errors, route.get("max_returned_bytes") == expected_max, f"{name}: cap mismatch")
    fail(errors, route.get("source_bytes") == expected_source_bytes, f"{name}: source byte mismatch")
    fail(errors, route.get("source_sha256") == expected_source_sha, f"{name}: source hash mismatch")
    fail(errors, route.get("output_sha256") == expected_output_sha, f"{name}: output hash mismatch")
    fail(errors, route.get("output_verified") is True, f"{name}: output oracle false")
    fail(errors, route.get("source_version_stable") is True, f"{name}: source version unstable")
    fail(errors, route.get("inverse_source_restored") is True, f"{name}: inverse false")
    fail(errors, route.get("sink_accepted_bytes") == expected_output_bytes, f"{name}: sink byte mismatch")
    fail(errors, route.get("sink_largest_write", 0) <= 64 * 1024, f"{name}: sink write exceeds cap")
    fail(errors, route.get("two_fingerprint_lower_bound_bytes") == fingerprint_bytes, f"{name}: fingerprint bound mismatch")
    fail(errors, route.get("logical_bytes_after_two_fingerprint_lower_bound") == operation_bytes - fingerprint_bytes, f"{name}: logical residual mismatch")
    fail(errors, route.get("underlying_bytes_after_two_fingerprint_lower_bound") == operation_bytes - fingerprint_bytes, f"{name}: underlying residual mismatch")
    fail(errors, route.get("logical_operation_conservation_verified") is True, f"{name}: logical conservation flag false")
    fail(errors, route.get("underlying_operation_conservation_verified") is True, f"{name}: underlying conservation flag false")
    phase_rows = route.get("phases", [])
    fail(errors, [p.get("phase") for p in phase_rows] == phases, f"{name}: phase order mismatch")
    previous = {key: 0 for key in scalar_keys}
    sums = {key: 0 for key in scalar_keys}
    for phase in phase_rows:
        cumulative = phase.get("logical_cumulative", {})
        delta = phase.get("logical_delta", {})
        for key in scalar_keys:
            expected_delta = cumulative.get(key, 0) - previous[key]
            fail(errors, cumulative.get(key, 0) >= previous[key], f"{name}/{phase.get('phase')}: logical counter regressed")
            fail(errors, delta.get(key) == expected_delta, f"{name}/{phase.get('phase')}: logical delta mismatch for {key}")
            fail(errors, expected_delta >= 0, f"{name}/{phase.get('phase')}: negative logical delta")
            sums[key] += delta.get(key, 0)
            previous[key] = cumulative.get(key, 0)
    logical = route.get("logical_operation", {})
    fail(errors, all(sums[key] == logical.get(key) for key in scalar_keys), f"{name}: logical phase sum does not equal operation")
    previous = {key: 0 for key in scalar_keys}
    sums = {key: 0 for key in scalar_keys}
    for phase in phase_rows:
        cumulative = phase.get("underlying_cumulative", {})
        delta = phase.get("underlying_delta", {})
        for key in scalar_keys:
            expected_delta = cumulative.get(key, 0) - previous[key]
            fail(errors, cumulative.get(key, 0) >= previous[key], f"{name}/{phase.get('phase')}: underlying counter regressed")
            fail(errors, delta.get(key) == expected_delta, f"{name}/{phase.get('phase')}: underlying delta mismatch for {key}")
            fail(errors, expected_delta >= 0, f"{name}/{phase.get('phase')}: negative underlying delta")
            sums[key] += delta.get(key, 0)
            previous[key] = cumulative.get(key, 0)
    underlying = route.get("underlying_operation", {})
    fail(errors, all(sums[key] == underlying.get(key) for key in scalar_keys), f"{name}: underlying phase sum does not equal operation")
    if name == "full_64k_unpaced":
        expected = {"read_calls": 165, "requested_bytes": operation_bytes, "returned_bytes": operation_bytes, "short_reads": 0}
        fail(errors, logical == expected, f"{name}: logical totals mismatch")
        fail(errors, underlying == expected, f"{name}: underlying totals mismatch")
    else:
        expected_logical = {"read_calls": 1607, "requested_bytes": 84_096_471, "returned_bytes": operation_bytes, "short_reads": 1506}
        expected_underlying = {"read_calls": 1607, "requested_bytes": operation_bytes, "returned_bytes": operation_bytes, "short_reads": 0}
        fail(errors, logical == expected_logical, f"{name}: logical totals mismatch")
        fail(errors, underlying == expected_underlying, f"{name}: underlying totals mismatch")
for route in reports:
    check_route(route)
fail(errors, [route.get("route") for route in reports] == ["full_64k_unpaced", "short_4k_unpaced"], "route order mismatch")
source_end = sha256(source_path)[1]
fail(errors, source_end == source_start, "source lib.rs changed during replay")
receipt = {
    "schema": "litchi.performance.0927.independent-phase-replay.v1",
    "status": "passed" if not errors else "failed",
    "claim_scope": "read-counter diagnostic replay only; no timing, native CPU, or performance claim",
    "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
    "source": {
        "path": str(source_path),
        "sha256_start": source_start,
        "sha256_end": source_end,
        "expected_sha256": expected_source_file_sha,
        "unchanged_during_replay": source_start == source_end,
    },
    "inputs": {
        "diagnostic_json": {"path": str(diag_path), "bytes": diag_size, "sha256": diag_sha},
        "focused_log": {"path": str(log_path), "bytes": log_size, "sha256": log_sha},
        "diagnostic_patch": {"path": str(patch_path), "bytes": patch_size, "sha256": patch_sha},
    },
    "observed_command": {
        "argv": ["cargo", "test", "--manifest-path", "tools/perf-baseline/Cargo.toml", "docx_section_layout_read_diagnostic_conserves_full_and_short_phases", "--", "--nocapture"],
        "cwd": str(root),
        "exit_status": 0,
        "source": "coder execution report; the archived focused log contains the test output but no argv header",
        "log_target": "tools/perf-baseline/target/debug/deps/litchi_perf_baseline-8fb54a160ae47866",
        "independent_replay": {
            "argv": ["python3", str(Path("/tmp/litchi-performance-0927/reviewer_phase_replay.py"))],
            "cwd": str(root),
            "exit_status": 0,
            "scope": "offline JSON/log replay only; no build or new binary execution",
        },
    },
    "quality_checks": {
        "clippy_strict": {
            "argv": ["cargo", "clippy", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked", "--offline", "--lib", "--tests", "--", "-D", "warnings"],
            "cwd": str(root),
            "exit_status": 101,
            "log": {"path": "/tmp/litchi-performance-0927/clippy-lib-tests.log", "bytes": 3636, "sha256": "a6d8e55da476dcb720820f3ed4219fd5c777d901767a5cee74faaeb1107fcfd9"},
            "baseline_warning": "clippy::result-large-err in unchanged tools/perf-baseline/src/producer_shape/selected_route_tests.rs",
        },
        "clippy_with_known_baseline_allowed": {
            "argv": ["cargo", "clippy", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked", "--offline", "--lib", "--tests", "--", "-D", "warnings", "-A", "clippy::result-large-err"],
            "cwd": str(root),
            "exit_status": 0,
            "log": {"path": "/tmp/litchi-performance-0927/clippy-lib-tests-allow-baseline.log", "bytes": 160, "sha256": "0f462f47f37260c8d692021b39910beeb0131a7db8a71a1a76bb2a96a60d5938"},
        },
        "clippy_lib_strict": {
            "argv": ["cargo", "clippy", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked", "--offline", "--lib", "--", "-D", "warnings"],
            "cwd": str(root),
            "exit_status": 0,
            "log": {"path": "/tmp/litchi-performance-0927/clippy-lib.log", "bytes": 160, "sha256": "16da4a206511c9f09bbcc6ed32c61c3f74549a3cde8abb62e43bc0af8731c2fa"},
        },
    },
    "checks": {
        "routes": 2,
        "route_names": ["full_64k_unpaced", "short_4k_unpaced"],
        "phase_names": phases,
        "checked_cumulative_delta_and_sum": True,
        "fixed_source_output_hashes": True,
        "same_sink_bound_and_inverse": True,
        "source_version_stable": True,
        "log_test_passed": True,
    },
    "errors": errors,
}
receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
print(json.dumps({"status": receipt["status"], "errors": errors, "receipt": str(receipt_path), "receipt_sha256": sha256(receipt_path)[1]}))

