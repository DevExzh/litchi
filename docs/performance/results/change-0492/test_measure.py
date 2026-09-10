#!/usr/bin/env python3
"""Fail-closed unit tests for the 0492 read-ahead evidence driver."""

from __future__ import annotations

import copy
import fcntl
import json
from pathlib import Path
import sys
import subprocess
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure as matrix  # noqa: E402


REVISION = "0123456789abcdef0123456789abcdef01234567"


def _counter(scope: str, requests: list[int], *, returned: list[int] | None = None,
             delayed: int = 0, paced: int = 0, delay_ns: int = 0) -> dict[str, object]:
    returned = requests if returned is None else returned
    counts = [0] * matrix.REQUEST_BUCKETS
    for length in requests:
        counts[matrix._bucket(length)] += 1
    return {
        "availability": "available",
        "logical_calls": len(requests),
        "max_request_bytes": max(requests) if requests else None,
        "min_request_bytes": min(requests) if requests else None,
        "request_size_counts": counts,
        "requested_bytes": sum(requests),
        "returned_bytes": sum(returned),
        "scope": scope,
        "short_reads": sum(a > b for a, b in zip(requests, returned)),
        "delayed_calls": delayed,
        "transfer_paced_calls": paced,
        "transfer_delay_ns": delay_ns,
    }


def _ranges(offsets: list[int], requests: list[int], returned: list[int] | None = None) -> list[dict[str, int]]:
    returned = requests if returned is None else returned
    return [{"offset": offset, "requested": request, "returned": result}
            for offset, request, result in zip(offsets, requests, returned)]


def _media(overlap: int = 0, calls: int = 2) -> dict[str, object]:
    return {
        "availability": "available",
        "media_range_count": len(matrix.MEDIA_RANGES),
        "observed_call_count": calls,
        "reason": None,
        "returned_overlap_bytes": overlap,
        "requested_overlap_bytes": overlap,
        "scope": matrix.EXPECTED_MEDIA_SCOPE,
        "status": "media_overlap_observed" if overlap else "proved_no_media_overlap",
    }


def _allocation() -> dict[str, object]:
    return {
        "allocated_bytes": 100,
        "allocation_calls": 4,
        "deallocated_bytes": 50,
        "deallocation_calls": 2,
        "failed_allocation_calls": 0,
        "live_bytes_after": 80,
        "live_bytes_before": 60,
        "peak_live_bytes_after": 120,
        "peak_live_bytes_before": 70,
        "reallocation_calls": 1,
        "region_peak_live_bytes": 100,
        "scope": matrix.EXPECTED_ALLOCATION_SCOPE_NAME,
        "status": matrix.EXPECTED_ALLOCATION_STATUS,
    }


def _row(index: int, arm: dict[str, object], role: str, latency: int = 100) -> dict[str, object]:
    candidate = arm.get("read_ahead_window_bytes") is not None
    logical_offsets = [offset for offset, _length in matrix.EXPECTED_LOGICAL_RANGES]
    logical_requests = [length for _offset, length in matrix.EXPECTED_LOGICAL_RANGES]
    if candidate:
        physical_requests = [22, 1327, matrix.WINDOW_BYTES]
        physical_offsets = [16_793_014, 16_791_709, 0]
        ahead = {
            "window_capacity": matrix.WINDOW_BYTES,
            "requests": 19,
            "hits": 16,
            "misses": 3,
            "fills": 3,
            "fill_requested_bytes": 5445,
            "fill_returned_bytes": 5445,
            "short_fills": 0,
            "max_fill_bytes": matrix.WINDOW_BYTES,
            "failures": 0,
        }
        # The metadata fill [0,4096) contributes 69 compressed media bytes.
        physical_media = _media(69, 3)
    else:
        physical_requests = logical_requests
        physical_offsets = logical_offsets
        ahead = None
        physical_media = _media(0, 19)
    logical = _counter(matrix.EXPECTED_LOGICAL_SCOPE, logical_requests)
    physical = _counter(matrix.EXPECTED_PHYSICAL_SCOPE, physical_requests)
    adapter = _counter(
        matrix.EXPECTED_RANGE_SCOPE,
        physical_requests,
        delayed=len(physical_requests) if arm["delay_us"] is not None else 0,
        paced=len(physical_requests) if arm["transfer_bytes_per_second"] is not None else 0,
        delay_ns=1000 if arm["delay_us"] else 0,
    )
    return {
        "cache": {
            "open_successful_loads": 0,
            "successful_loads": 1,
            "failed_loads": 0,
            "retained_bytes": 1424,
            "retained_entries": 1,
            "budget_managed": False,
        },
        "actual_text_bytes": matrix.CORPUS["expected_text_bytes"],
        "actual_text_sha256": matrix.CORPUS["expected_text_sha256"],
        "actual_text_verified": True,
        **({"allocation": _allocation()} if role == "allocator" else {}),
        "latency_ns": latency,
        "reads": {
            **({"read_ahead": ahead} if candidate else {}),
            "logical_wrapper": logical,
            "logical_ranges": _ranges(logical_offsets, logical_requests),
            "physical_ranges": _ranges(physical_offsets, physical_requests),
            "logical_media_range_proof": _media(0, len(logical_requests)),
            "wrapper": physical,
            "range_adapter": adapter,
            "media_range_proof": physical_media,
        },
        "sample_index": index,
        "source_version_after": {"id": 188, "revision": 0},
        "source_version_before": {"id": 188, "revision": 0},
        "source_version_unchanged": True,
    }


def _report(arm_name: str = "baseline-65536-0us", role: str = "normal",
            samples: int = 1, warmups: int = 0, latency: int = 100) -> dict[str, object]:
    arm = matrix.ARM_BY_NAME[arm_name]
    provider = {
        "delay_us": arm["delay_us"],
        "file_scope": matrix.EXPECTED_FILE_SCOPE,
        "max_range_bytes": arm["max_range_bytes"],
        "provider": arm["provider"],
        "read_counter_scope": matrix.GENERAL_READ_SCOPE,
        "source_construction": matrix.EXPECTED_SOURCE_CONSTRUCTION,
        "transfer_bytes_per_second": arm["transfer_bytes_per_second"],
        "transfer_delay_policy": arm["transfer_delay_policy"],
    }
    if arm.get("read_ahead_window_bytes") is not None:
        provider["read_ahead_window_bytes"] = arm["read_ahead_window_bytes"]
    return {
        "allocation_scope": matrix.EXPECTED_ALLOCATION_SCOPE,
        "corpus": dict(matrix.CORPUS),
        "limits": dict(matrix.LIMITS),
        "media_ranges": [{"start": start, "end": end} for start, end in matrix.MEDIA_RANGES],
        "provider": provider,
        "provider_scope": matrix.EXPECTED_PROVIDER_SCOPE,
        "requested_source_revision": REVISION,
        "rows": [_row(index + warmups, arm, role, latency + index)
                 for index in range(samples)],
        "schema": matrix.SCHEMA,
        "samples": samples,
        "setup_scope": matrix.EXPECTED_SETUP_SCOPE,
        "source_bytes": matrix.CORPUS["archive_bytes"],
        "source_sha256": matrix.CORPUS["archive_sha256"],
        "timing_scope": matrix.EXPECTED_TIMING_SCOPE,
        "warmup": warmups,
    }


def _binary() -> dict[str, object]:
    return {"bytes": 1, "executable": True, "path": "/bin/true", "sha256": "a" * 64}


def _entry(role: str, arm_name: str, repeat: int, latency: int) -> dict[str, object]:
    report = _report(arm_name, role, samples=1, warmups=0, latency=latency)
    spec = {
        "kind": "formal",
        "repeat": repeat,
        "role": role,
        "arm": arm_name,
        "provider": matrix.ARM_BY_NAME[arm_name]["provider"],
        "samples": 1,
        "warmups": 0,
        "label": f"r{repeat}-{role}-{arm_name}",
    }
    return {
        "spec": spec,
        "terminal_sha256": "b" * 64,
        "report_sha256": "c" * 64,
        "report": report,
        "resource": {"maximum_resident_set_size_(kbytes)": 10},
    }


class ReadAheadMatrixTests(unittest.TestCase):
    def test_inventory_is_sixteen_processes_and_reverses_order(self) -> None:
        runs = matrix.formal_inventory()
        self.assertEqual(len(runs), 16)
        self.assertEqual([item["arm"] for item in runs[:4]], [arm["name"] for arm in matrix.ARMS])
        self.assertEqual([item["role"] for item in runs[:4]], ["normal"] * 4)
        self.assertEqual([item["arm"] for item in runs[4:8]], [arm["name"] for arm in matrix.ARMS])
        self.assertEqual([item["role"] for item in runs[4:8]], ["allocator"] * 4)
        self.assertEqual([item["arm"] for item in runs[8:12]], [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual([item["role"] for item in runs[8:12]], ["allocator"] * 4)
        self.assertEqual([item["arm"] for item in runs[12:]], [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual([item["role"] for item in runs[12:]], ["normal"] * 4)
        self.assertEqual(len({item["label"] for item in runs}), 16)

    def test_frozen_arms_bind_controls_and_candidate_window(self) -> None:
        zero = matrix.ARM_BY_NAME["candidate-4096-0us"]
        delayed = matrix.ARM_BY_NAME["candidate-4096-1000us-104857600bps-minimum-service"]
        self.assertEqual(zero["read_ahead_window_bytes"], 4096)
        self.assertEqual(zero["max_range_bytes"], 65536)
        self.assertEqual(delayed["delay_us"], 1000)
        self.assertEqual(delayed["transfer_bytes_per_second"], 104857600)

    def test_report_validates_baseline_and_candidate_v2_shapes(self) -> None:
        for arm_name in ("baseline-65536-0us", "candidate-4096-0us"):
            with self.subTest(arm_name=arm_name), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "report.json"
                value = _report(arm_name)
                path.write_text(json.dumps(value), encoding="utf-8")
                checked = matrix.validate_report(path, role="normal", arm_name=arm_name,
                                                 binary=_binary(), samples=1, warmups=0,
                                                 source_revision=REVISION)
                self.assertEqual(checked["schema"], matrix.SCHEMA)

    def test_source_and_range_conservation_is_strict(self) -> None:
        value = _report("candidate-4096-0us")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                   binary=_binary(), samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["physical_ranges"][0]["requested"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

            # Keep all totals and the request histogram unchanged while moving
            # one logical range.  The replay must match the pinned sequence,
            # rather than merely conserve aggregate bytes.
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["logical_ranges"][0]["offset"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

            # A summary extrema field is part of the receipt and must be
            # recomputed from the range trace on every validation.
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["wrapper"]["max_request_bytes"] -= 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

            # Every read-ahead counter is consumed by the conservation checks;
            # an unexplained failure must fail closed even when the other
            # fields still look internally consistent.
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["read_ahead"]["failures"] = 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_media_range_count_is_pinned_to_the_corpus(self) -> None:
        value = _report("candidate-4096-0us")
        value["rows"][0]["reads"]["media_range_proof"]["media_range_count"] -= 1
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_candidate_fill_conservation_and_bound_are_strict(self) -> None:
        value = _report("candidate-4096-0us")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                   binary=_binary(), samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["read_ahead"]["max_fill_bytes"] = matrix.WINDOW_BYTES + 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_unmanaged_cache_diagnostic_rejects_managed_flag(self) -> None:
        value = _report("candidate-4096-0us")
        value["rows"][0]["cache"]["budget_managed"] = True
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_report_rejects_missing_or_changed_oracle(self) -> None:
        value = _report()
        value["rows"][0].pop("actual_text_sha256")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="baseline-65536-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_report_rejects_source_version_mutation_and_bad_indices(self) -> None:
        value = _report(samples=2)
        value["rows"][1]["sample_index"] = 0
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="baseline-65536-0us",
                                       binary=_binary(), samples=2, warmups=0)
            value = _report()
            value["rows"][0]["source_version_after"]["revision"] = 1
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="baseline-65536-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_media_overlap_is_measured_without_claiming_zero(self) -> None:
        value = _report("candidate-4096-0us")
        value["rows"][0]["reads"]["media_range_proof"]["requested_overlap_bytes"] = 118
        value["rows"][0]["reads"]["media_range_proof"]["returned_overlap_bytes"] = 118
        value["rows"][0]["reads"]["media_range_proof"]["status"] = "media_overlap_observed"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            # Summary-only overlap is intentionally rejected because the trace
            # must reconcile to the proof; this mutation has no corresponding
            # physical range and therefore fails closed.
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="candidate-4096-0us",
                                       binary=_binary(), samples=1, warmups=0)

    def test_bootstrap_is_reproducible_and_has_tail_quantiles(self) -> None:
        values = [1, 2, 3, 4, 10, 11]
        first = matrix._bootstrap_median_ci(values)
        second = matrix._bootstrap_median_ci(values)
        self.assertEqual(first, second)
        self.assertEqual(first["confidence"], 0.95)
        self.assertIn("ci_low", first)
        self.assertIn("ci_high", first)

    def test_analyze_exposes_comparisons_and_adverse_tails(self) -> None:
        entries = []
        for role in matrix.ROLES:
            for arm in matrix.ARMS:
                entries.append(_entry(role, arm["name"], 1, 100))
        builds = {
            role: {"receipt_sha256": "d" * 64,
                   "binary": {"path": "/bin/true", "bytes": 1, "sha256": "a" * 64},
                   "source": {"path": "source.json", "files": 1, "sha256": "e" * 64},
                   "gate": {"path": "gate.json", "sha256": "f" * 64}}
            for role in matrix.ROLES
        }
        summary = matrix.analyze_data(entries, builds, pilot=True)
        self.assertEqual(summary["inventory"]["formal_processes"], 8)
        self.assertEqual(summary["inventory"]["measured_samples"], 8)
        self.assertEqual(len(summary["comparisons"]), 4)
        self.assertIn("bootstrap_median_ci", summary["rows"][0])
        self.assertTrue(summary["adverse_rows_over_5_percent"])
        self.assertTrue(any(item["metric"] == "physical_read_requested_bytes"
                            and item["percentile"] == "p99"
                            for item in summary["adverse_rows_over_5_percent"]))

    def test_private_cleanup_refuses_unknown_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "provider" / "attempt" / "label"
            tmp = root / "tmp"
            tmp.mkdir(parents=True)
            (tmp / "unexpected").write_bytes(b"owned-test")
            with patch.object(matrix, "TEMP", Path(directory)):
                self.assertEqual(matrix._cleanup_private(root, tmp)["status"], "failed")
                self.assertTrue((tmp / "unexpected").is_file())
                (tmp / "unexpected").unlink()
                self.assertEqual(matrix._cleanup_private(root, tmp)["status"], "pass")
                self.assertFalse(root.exists())

    def test_command_binds_trace_and_candidate_flags(self) -> None:
        binary = {"path": "/bin/true"}
        command = matrix._command(matrix.formal_inventory()[1], binary,
                                  Path("report.json"), Path("time.txt"), REVISION)
        self.assertIn("--trace-ranges", command)
        self.assertIn("--read-ahead", command)
        self.assertIn("4096", command)
        delayed = matrix.formal_inventory()[3]
        command = matrix._command(delayed, binary, Path("report.json"),
                                  Path("time.txt"), REVISION)
        self.assertIn("--transfer-bytes-per-second", command)
        self.assertIn("104857600", command)
        self.assertIn("minimum-service", command)

    def test_cpu_lock_is_held_during_capture_callback_and_released_afterward(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            lock_path = Path(directory) / "cpu.lock"
            observed: list[bool] = []

            def callback() -> str:
                with lock_path.open("a+") as contender:
                    try:
                        fcntl.flock(contender.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    except BlockingIOError:
                        observed.append(True)
                    else:  # pragma: no cover - proves the lock was not held
                        observed.append(False)
                        fcntl.flock(contender.fileno(), fcntl.LOCK_UN)
                return "done"

            with patch.object(matrix, "CPU_LOCK", str(lock_path)):
                self.assertEqual(matrix._with_cpu_lock(callback), "done")
                with lock_path.open("a+") as released:
                    fcntl.flock(released.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    fcntl.flock(released.fileno(), fcntl.LOCK_UN)
            self.assertEqual(observed, [True])

    def test_timeout_kills_the_process_group_and_preserves_terminal_state(self) -> None:
        process = subprocess.Popen(["sleep", "30"], start_new_session=True)
        try:
            termination = matrix._kill_group(process)
            self.assertEqual(termination, "SIGTERM")
            self.assertIsNotNone(process.returncode)
        finally:
            if process.poll() is None:
                process.kill()
                process.wait()


if __name__ == "__main__":
    unittest.main()
