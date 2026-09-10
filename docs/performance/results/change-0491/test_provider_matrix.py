#!/usr/bin/env python3
"""Fail-closed tests for :mod:`provider_matrix`."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import provider_matrix as matrix  # noqa: E402


def _counter(available: bool, *, range_counter: bool = False) -> dict[str, object]:
    scope = matrix.EXPECTED_RANGE_SCOPE if range_counter else (matrix.EXPECTED_READ_SCOPE if available else matrix.GENERAL_READ_SCOPE)
    if not available:
        return {
            "availability": "unavailable",
            "logical_calls": None,
            "max_request_bytes": None,
            "min_request_bytes": None,
            "request_size_counts": None,
            "requested_bytes": None,
            "returned_bytes": None,
            "scope": scope,
            "short_reads": None,
            "delayed_calls": None,
            "transfer_paced_calls": None,
            "transfer_delay_ns": None,
        }
    return {
        "availability": "available",
        "logical_calls": 2,
        "max_request_bytes": 8,
        "min_request_bytes": 4,
        "request_size_counts": [1, 1] + [0] * 16,
        "requested_bytes": 12,
        "returned_bytes": 12,
        "scope": scope,
        "short_reads": 0,
        "delayed_calls": 2 if range_counter else 0,
        "transfer_paced_calls": 1 if range_counter else 0,
        "transfer_delay_ns": 100 if range_counter else 0,
    }


def _media(available: bool) -> dict[str, object]:
    return {
        "availability": "available" if available else "unavailable",
        "media_range_count": 8,
        "observed_call_count": 2 if available else None,
        "reason": None if available else "provider has no offset instrumentation",
        "returned_overlap_bytes": 0 if available else None,
        "requested_overlap_bytes": 0 if available else None,
        "scope": matrix.EXPECTED_MEDIA_SCOPE,
        "status": "proved_no_media_overlap" if available else "unavailable",
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
    provider = str(arm["provider"])
    counted = provider in {"file", "instrumented-bytes", "range"}
    return {
        "actual_text_bytes": matrix.CORPUS["expected_text_bytes"],
        "actual_text_sha256": matrix.CORPUS["expected_text_sha256"],
        "actual_text_verified": True,
        **({"allocation": _allocation()} if role == "allocator" else {}),
        "latency_ns": latency,
        "reads": {
            "media_range_proof": _media(counted),
            "range_adapter": _counter(True, range_counter=True) if provider == "range" else None,
            "wrapper": _counter(counted),
        },
        "sample_index": index,
        "source_version_after": {"id": 188, "revision": 0},
        "source_version_before": {"id": 188, "revision": 0},
        "source_version_unchanged": True,
    }


def _report(arm_name: str = "bytes", role: str = "normal", samples: int = 1,
            warmups: int = 0, latency: int = 100) -> dict[str, object]:
    arm = matrix.ARM_BY_NAME[arm_name]
    return {
        "allocation_scope": matrix.EXPECTED_ALLOCATION_SCOPE,
        "corpus": dict(matrix.CORPUS),
        "limits": dict(matrix.LIMITS),
        "provider": {
            "delay_us": arm["delay_us"],
            "file_scope": matrix.EXPECTED_FILE_SCOPE,
            "max_range_bytes": arm["max_range_bytes"],
            "provider": arm["provider"],
            "read_counter_scope": matrix.GENERAL_READ_SCOPE,
            "source_construction": matrix.EXPECTED_SOURCE_CONSTRUCTION,
            "transfer_bytes_per_second": arm["transfer_bytes_per_second"],
            "transfer_delay_policy": arm["transfer_delay_policy"],
        },
        "provider_scope": matrix.EXPECTED_PROVIDER_SCOPE,
        "requested_source_revision": "0123456789abcdef0123456789abcdef01234567",
        "rows": [_row(index + warmups, arm, role, latency + index) for index in range(samples)],
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


class ProviderMatrixTests(unittest.TestCase):
    def test_inventory_has_twenty_processes_and_reverses_role_and_arm_order(self) -> None:
        runs = matrix.formal_inventory()
        self.assertEqual(len(runs), 20)
        self.assertEqual([item["arm"] for item in runs[:5]], [arm["name"] for arm in matrix.ARMS])
        self.assertEqual([item["role"] for item in runs[:5]], ["normal"] * 5)
        self.assertEqual([item["arm"] for item in runs[5:10]], [arm["name"] for arm in matrix.ARMS])
        self.assertEqual([item["role"] for item in runs[5:10]], ["allocator"] * 5)
        self.assertEqual([item["arm"] for item in runs[10:15]], [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual([item["role"] for item in runs[10:15]], ["allocator"] * 5)
        self.assertEqual([item["arm"] for item in runs[15:]], [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual([item["role"] for item in runs[15:]], ["normal"] * 5)
        self.assertEqual(len({item["label"] for item in runs}), 20)

    def test_frozen_arms_bind_range_delay_and_minimum_service(self) -> None:
        zero = matrix.ARM_BY_NAME["range-64-0us"]
        delayed = matrix.ARM_BY_NAME["range-65536-1000us-104857600bps-minimum-service"]
        self.assertEqual(zero["max_range_bytes"], 64)
        self.assertEqual(zero["delay_us"], 0)
        self.assertIsNone(zero["transfer_bytes_per_second"])
        self.assertEqual(delayed["delay_us"], 1000)
        self.assertEqual(delayed["transfer_bytes_per_second"], 104857600)
        self.assertEqual(delayed["transfer_delay_policy"], "minimum-service")

    def test_report_validates_all_provider_identity_fields(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            value = _report("range-65536-1000us-104857600bps-minimum-service")
            path.write_text(json.dumps(value), encoding="utf-8")
            checked = matrix.validate_report(
                path,
                role="normal",
                arm_name="range-65536-1000us-104857600bps-minimum-service",
                binary=_binary(),
                samples=1,
                warmups=0,
                source_revision=value["requested_source_revision"],
            )
            self.assertEqual(checked["corpus"], matrix.CORPUS)

    def test_short_read_arm_requires_positive_capped_observation(self) -> None:
        value = _report("range-64-0us")
        adapter = value["rows"][0]["reads"]["range_adapter"]
        adapter.update({
            "max_request_bytes": 128,
            "requested_bytes": 128,
            "returned_bytes": 64,
            "short_reads": 1,
        })
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="range-64-0us", binary=_binary(), samples=1, warmups=0)
            adapter["short_reads"] = 0
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="range-64-0us", binary=_binary(), samples=1, warmups=0)

    def test_report_rejects_missing_actual_text_oracle(self) -> None:
        value = _report()
        value["rows"][0].pop("actual_text_sha256")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="bytes", binary=_binary(), samples=1, warmups=0)

    def test_report_rejects_corpus_drift_and_extra_schema_fields(self) -> None:
        for mutation in ("corpus", "extra"):
            value = _report()
            if mutation == "corpus":
                value["corpus"]["archive_sha256"] = "f" * 64
            else:
                value["unexpected"] = True
            with self.subTest(mutation=mutation), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "report.json"
                path.write_text(json.dumps(value), encoding="utf-8")
                with self.assertRaises(matrix.ProviderMatrixError):
                    matrix.validate_report(path, role="normal", arm_name="bytes", binary=_binary(), samples=1, warmups=0)

    def test_allocator_role_requires_measured_allocation_record(self) -> None:
        value = _report("bytes", role="allocator")
        value["rows"][0].pop("allocation")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="allocator", arm_name="bytes", binary=_binary(), samples=1, warmups=0)

    def test_source_version_change_is_rejected(self) -> None:
        value = _report()
        value["rows"][0]["source_version_after"]["revision"] = 1
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="bytes", binary=_binary(), samples=1, warmups=0)

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

    def test_source_binding_rejects_manifest_content_drift(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest_path = Path(directory) / "sources.json"
            manifest = {"src/lib.rs": "a" * 64}
            encoded = (json.dumps(manifest, indent=2, sort_keys=True) + "\n").encode()
            manifest_path.write_bytes(encoded)
            binding = {
                "files": 1,
                "path": str(manifest_path),
                "sha256": hashlib.sha256(encoded).hexdigest(),
            }
            matrix._source_binding(binding, manifest_path, "source")
            manifest_path.write_text("{}\n", encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._source_binding(binding, manifest_path, "source")

    def test_gate_binding_rejects_nonzero_exit(self) -> None:
        source_gate = Path(__file__).resolve().parent / "validation" / "build-normal-pilot1.json"
        if not source_gate.is_file():
            self.skipTest("retained pilot gate receipt is not present")
        value = json.loads(source_gate.read_text(encoding="utf-8"))
        value["exit_code"] = 1
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            gate_path = root / "gate.json"
            value["artifacts"] = {
                "gate.stdout": value["artifacts"][next(key for key in value["artifacts"] if key.endswith(".stdout"))],
                "gate.stderr": value["artifacts"][next(key for key in value["artifacts"] if key.endswith(".stderr"))],
            }
            for suffix in ("stdout", "stderr"):
                original = source_gate.with_suffix(f".{suffix}")
                (root / f"gate.{suffix}").write_bytes(original.read_bytes())
            gate_path.write_text(json.dumps(value), encoding="utf-8")
            binding = {"path": str(gate_path), "sha256": matrix.sha(gate_path)}
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._gate_binding(binding, root / "build.json")

    def test_capture_artifact_binding_rejects_escape_and_empty_resource(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            capture = Path(directory) / "capture"
            capture.mkdir()
            outside = Path(directory) / "outside-report.json"
            outside.write_text("{}", encoding="utf-8")
            for name in ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json"):
                (capture / name).write_bytes(b"x")
            artifacts = {
                name: {"path": str(capture / name), "bytes": 1, "sha256": matrix.sha(capture / name)}
                for name in ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json")
            }
            artifacts["report.json"] = {"path": str(outside), "bytes": outside.stat().st_size,
                                         "sha256": matrix.sha(outside)}
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._check_report_artifacts(capture, {"artifacts": artifacts})

    def test_repeat_variance_is_descriptive_and_retains_quantiles(self) -> None:
        entries = []
        for role in matrix.ROLES:
            for arm in matrix.ARMS:
                for repeat, latency in ((1, 100), (2, 110)):
                    entries.append(_entry(role, arm["name"], repeat, latency))
        builds = {
            role: {"receipt_sha256": "d" * 64, "binary": {"path": "/bin/true", "bytes": 1, "sha256": "a" * 64},
                   "source": {"path": "source.json", "files": 1, "sha256": "e" * 64},
                   "gate": {"path": "gate.json", "sha256": "f" * 64}}
            for role in matrix.ROLES
        }
        summary = matrix.analyze_data(entries, builds)
        self.assertEqual(summary["inventory"]["formal_processes"], 20)
        self.assertEqual(summary["inventory"]["measured_samples"], 20)
        self.assertEqual(len(summary["repeat_variance"]), 10)
        pair = summary["repeat_variance"][0]
        self.assertIn("p50", pair["metrics"]["latency_ns"]["relative_percent"])
        self.assertTrue(pair["metrics"]["latency_ns"]["flag_over_5_percent"])
        self.assertIn("no optimization regression claim", pair["scope"])

    def test_command_binds_all_delayed_range_parameters(self) -> None:
        spec = matrix.formal_inventory()[4]
        command = matrix._command(spec, {"path": "/bin/true"}, Path("report.json"), Path("time.txt"), "0" * 40)
        self.assertIn("--max-range", command)
        self.assertIn("65536", command)
        self.assertIn("--delay-us", command)
        self.assertIn("1000", command)
        self.assertIn("--transfer-bytes-per-second", command)
        self.assertIn("104857600", command)
        self.assertIn("minimum-service", command)


class RetainedReportTests(unittest.TestCase):
    def test_real_pre_freeze_reports_match_schema(self):
        root = Path(__file__).resolve().parent
        binary = json.loads((root / "build-normal-integrated1.json").read_text())["binary"]
        for name, arm in [("bytes", "bytes"), ("file", "file"), ("instrumented-bytes", "instrumented-bytes"), ("delayed", matrix.ARMS[-1]["name"])]:
            with self.subTest(name=name):
                report = root / "pilot-provider1" / f"{name}.json"
                value = json.loads(report.read_text())
                matrix.validate_report(report, role="normal", arm_name=arm, binary=binary,
                                       samples=1, warmups=0, source_revision=value["requested_source_revision"])

    def test_actual_delayed_adapter_cannot_omit_service_or_exceed_calls(self):
        arm = matrix.ARMS[-1]
        source = Path(__file__).parent / "captures/provider-pilot-final6" / ("pilot-r1-normal-" + arm["name"]) / "report.json"
        original = json.loads(source.read_text())["rows"][0]
        matrix._check_sample(original, original["sample_index"], arm, "normal", "test")
        for field, value in [("delayed_calls", 0), ("transfer_paced_calls", 0), ("transfer_delay_ns", 0),
                             ("short_reads", 999), ("min_request_bytes", 99999)]:
            altered = copy.deepcopy(original)
            altered["reads"]["range_adapter"][field] = value
            with self.subTest(field=field), self.assertRaises(matrix.ProviderMatrixError):
                matrix._check_sample(altered, original["sample_index"], arm, "normal", "test")

    def test_chronology_rejects_overlap_and_reordering(self):
        def entry(start, end):
            return {"started": {"started_utc": f"2026-09-09T00:00:{start:02}+00:00"},
                    "terminal": {"finished_utc": f"2026-09-09T00:00:{end:02}+00:00"}}
        matrix._check_chronology([entry(1, 2), entry(3, 4)])
        for entries in [[entry(1, 3), entry(2, 4)], [entry(3, 4), entry(1, 2)]]:
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._check_chronology(entries)


if __name__ == "__main__":
    unittest.main()
