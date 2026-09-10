#!/usr/bin/env python3
"""Focused fail-closed tests for :mod:`cold_matrix`."""

from __future__ import annotations

from contextlib import contextmanager
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import cold_matrix as matrix  # noqa: E402


@contextmanager
def _frozen_protocol(directory: Path):
    previous = matrix.PROTOCOL_FILE
    matrix.PROTOCOL_FILE = directory / "cold-protocol.json"
    try:
        matrix._freeze_protocol()
        yield matrix.PROTOCOL_FILE
    finally:
        matrix.PROTOCOL_FILE = previous


def _stats(samples: list[int]) -> dict[str, object]:
    ordered = sorted(samples)
    return {
        "unit": "ns",
        "samples": samples,
        "min": ordered[0],
        "p50": (ordered[14] + ordered[15]) // 2,
        "p95": ordered[28],
        "p99": ordered[29],
        "max": ordered[-1],
        "mean": sum(samples) / len(samples),
    }


def _replay() -> dict[str, object]:
    value = {
        "source_bytes": matrix.CORPUS["archive_bytes"],
        "source_sha256": matrix.CORPUS["archive_sha256"],
        "paragraph_count": matrix.EXPECTED_PARAGRAPHS,
        "operation": "full_text",
        "classification": matrix.EXPECTED_REPLAY_CLASSIFICATION,
        "semantic_sha256": matrix.EXPECTED_TEXT_SHA256,
        "materializations": 0,
        "preparation_main_payload_fully_covered": True,
    }
    for phase, sizes in [("open", [22, 46]), ("preparation", [1424]), ("query", [])]:
        value[f"{phase}_read_return_sizes"] = sizes
        value[f"{phase}_read_calls"] = len(sizes)
        value[f"{phase}_read_bytes"] = sum(sizes)
        for category in ("main", "media", "unselected", "core"):
            for suffix in ("overlap_bytes", "covered_bytes"):
                value[f"{phase}_{category}_payload_{suffix}"] = 1424 if (phase, category) == ("preparation", "main") else 0
    return value


def _sample(index: int, state: str = "warm") -> dict[str, object]:
    replay = _replay()
    return {
        "sample_index": index,
        "cache_state": state,
        "elapsed_ns": 100 + index,
        "parent_wall_ns": 200 + index,
        "cold_advice": "not_requested" if state == "warm" else "requested",
        "cold_verified": None,
        "docx_source_replay": replay,
        "docx_timed_full_text_sha256": replay["semantic_sha256"],
        "docx_timed_full_text_bytes": matrix.EXPECTED_TEXT_BYTES,
        "docx_timed_full_text_timing_scope": matrix.EXPECTED_TEXT_SCOPE,
        "logical_read_calls": 1,
        "logical_read_bytes": 2,
    }


def _report(state: str = "warm") -> dict[str, object]:
    samples = [_sample(index, state) for index in range(30)]
    return {
        "schema_version": 1,
        "tool": {
            "binary": "litchi-perf-baseline",
            "instrumentation": "none",
        },
        "configuration": {
            "cases": [matrix.CASE],
            "samples_per_case": 30,
            "warmup_iterations_per_case": 3,
            "filesystem_cache_states": [state],
            "filesystem_fresh_child_per_sample": True,
            "filesystem_process_isolated": True,
            "filesystem_root_selected": True,
        },
        "results": [{
            "case": matrix.CASE,
            "cache_state": state,
            "corpus": dict(matrix.CORPUS),
            "elapsed_ns": _stats([100 + index for index in range(30)]),
        }],
        "filesystem_evidence": [{
            "case": matrix.CASE,
            "corpus": dict(matrix.CORPUS),
            "warmup_iterations": 3,
            "sample_count": 30,
            "cache_states": [state],
            "fresh_child_per_sample": True,
            "samples": samples,
        }],
    }


def _eligible_proof() -> dict[str, object]:
    return {
        "status": "eligible",
        "filesystem_magic": 0xEF53,
        "page_size_bytes": 4096,
        "source_bytes": 16_793_600,
        "source_pages": 4100,
        "aligned_source_bytes": 16_793_600,
        "aligned_source_sha256": matrix.ALIGNMENT_ORACLE["aligned_source_sha256"],
        "fsync_completed": True,
        "advice": matrix.EXPECTED_COLD_ADVICE,
        "fincore_size_bytes": 16_793_600,
        "resident_bytes": 0,
        "dirty_bytes": 0,
        "writeback_bytes": 0,
        "fincore_tool": "fincore",
        "fincore_sha256": "2" * 64,
        "fincore_version": "util-linux",
        "fincore_stderr_sha256": "3" * 64,
        "fincore_stderr_bytes": 0,
        "fincore_version_stderr_sha256": "4" * 64,
        "fincore_version_stderr_bytes": 0,
        "fincore_method": matrix.EXPECTED_FINCORE_METHOD,
        "fincore_fallback": matrix.EXPECTED_FINCORE_FALLBACK,
        "read_bytes_before": 10,
        "read_bytes_after": 4106,
        "read_bytes_delta": 4096,
    }


def _write_build_fixture(directory: Path, *, gate_exit: int = 0,
                         source_sha256: str | None = None,
                         gate_command: list[str] | None = None) -> tuple[Path, dict[str, object]]:
    """Create a content-bound normal build/gate pair without executing it."""

    manifest_path = directory / "manifest.json"
    manifest_path.write_text(json.dumps({"src/lib.rs": "0" * 64}, sort_keys=True) + "\n", encoding="utf-8")
    manifest_hash = matrix.sha(manifest_path)
    binary_path = directory / "litchi-perf-baseline"
    binary_path.write_bytes(b"fixture executable\n")
    binary_path.chmod(0o755)
    command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
    ]
    source = {
        "path": manifest_path.name,
        "files": 1,
        "sha256": source_sha256 or manifest_hash,
    }
    (directory / "gate.stdout").write_bytes(b"")
    (directory / "gate.stderr").write_bytes(b"")
    gate = {
        "schema": "docx-stream-append-gate-v1",
        "label": "gate",
        "argv": gate_command or command,
        "cwd": str(matrix.REPO),
        "environment": matrix._environment(),
        "driver_sha256": matrix._json_hash(matrix.ROOT / "gate.py"),
        "common_sha256": matrix._json_hash(matrix.ROOT / "support.py"),
        "started_utc": "2026-09-09T00:00:00+00:00",
        "finished_utc": "2026-09-09T00:00:01+00:00",
        "exit_code": gate_exit,
        "source_before": source,
        "source_after": source,
        "source_unchanged": True,
        "artifacts": {
            "gate.stdout": matrix.meta(directory / "gate.stdout"),
            "gate.stderr": matrix.meta(directory / "gate.stderr"),
        },
    }
    gate_path = directory / "gate.json"
    gate_path.write_text(json.dumps(gate, sort_keys=True) + "\n", encoding="utf-8")
    build = {
        "role": "normal",
        "binary": {**matrix.meta(binary_path), "path": str(binary_path), "executable": True},
        "command": command,
        "environment": matrix._environment(),
        "source_before": source,
        "source_after": source,
        "source_unchanged": True,
        "gate": {"path": str(gate_path), **matrix.meta(gate_path)},
    }
    build_path = directory / "build-normal.json"
    build_path.write_text(json.dumps(build, sort_keys=True) + "\n", encoding="utf-8")
    return build_path, build


class ColdMatrixTests(unittest.TestCase):
    def test_inventory_is_bounded_and_reverses_roles(self) -> None:
        runs = matrix.formal_inventory()
        self.assertEqual(len(runs), 16)
        self.assertEqual(
            [run["role"] for run in runs[:8]],
            ["normal"] * 4 + ["allocator"] * 4,
        )
        self.assertEqual(
            [run["role"] for run in runs[8:]],
            ["allocator"] * 4 + ["normal"] * 4,
        )
        self.assertEqual(len({run["label"] for run in runs}), len(runs))

    def test_corpus_hash_drift_is_rejected(self) -> None:
        drifted = dict(matrix.CORPUS)
        drifted["archive_sha256"] = "f" * 64
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._manifest_matches(drifted, "corpus")

    def test_build_gate_must_record_success(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            build_path, build = _write_build_fixture(Path(directory), gate_exit=1)
            with self.assertRaises(matrix.ColdMatrixError):
                matrix._binary_from_build(build, "normal", build_path)

    def test_build_source_manifest_content_hash_is_bound(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            build_path, build = _write_build_fixture(Path(directory), source_sha256="f" * 64)
            with self.assertRaises(matrix.ColdMatrixError):
                matrix._binary_from_build(build, "normal", build_path)

    def test_build_gate_command_must_match_build_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            build_path, build = _write_build_fixture(
                Path(directory), gate_command=["cargo", "build", "--release"]
            )
            with self.assertRaises(matrix.ColdMatrixError):
                matrix._binary_from_build(build, "normal", build_path)

    def test_capture_labels_reject_extra_on_disk_directory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "expected").mkdir()
            (root / "unexpected").mkdir()
            with self.assertRaises(matrix.ColdMatrixError):
                matrix._check_label_set(root, ["expected"])

    def test_capture_artifact_path_must_stay_inside_label_directory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inside = root / "stdout.txt"
            outside = root.parent / "escaped.stdout"
            inside.write_bytes(b"inside")
            outside.write_bytes(b"outside")
            value = {"path": str(outside), **matrix.meta(outside)}
            with self.assertRaises(matrix.ColdMatrixError):
                matrix._confined_artifact(root, "stdout.txt", value)

    def test_capture_header_rejects_tampered_argv(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            with _frozen_protocol(Path(directory)):
                build_path, build_record = _write_build_fixture(Path(directory))
                build = matrix._binary_from_build(build_record, "normal", build_path)
                spec = {**matrix.formal_inventory(pilot=True)[0], "attempt": "unit"}
                capture = matrix.ROOT / "captures" / "unit" / spec["label"]
                expected_binary = build["binary"]
                expected = {
                    "run": spec,
                    "build": build["binding"],
                    "binary": expected_binary,
                    "source": build["source"],
                    "tool": matrix._tool_identity("normal"),
                    "machine": matrix._machine_binding(),
                    "protocol": matrix._protocol_binding(),
                    "argv": matrix._command(
                        spec, expected_binary, capture / "report.json",
                        capture / "corpus-manifest.json",
                        matrix._run_root("unit", spec["label"]) / "filesystem",
                    ),
                    "cwd": str(matrix.REPO),
                    "environment": matrix._environment(),
                    "driver_sha256": matrix._json_hash(Path(matrix.__file__)),
                    "support_sha256": matrix._json_hash(matrix.ROOT / "support.py"),
                }
                started = {"schema": matrix.SCHEMA, "version": 1, "status": "running", **expected,
                           "started_utc": "2026-09-09T00:00:00+00:00",
                           "process_group": {"start_new_session": True}}
                terminal = {**started, "status": "pass", "exit_code": 0, "timed_out": False,
                            "termination": None, "missing_artifacts": [],
                            "finished_utc": "2026-09-09T00:00:01+00:00",
                            "replay_root": str(matrix._run_root("unit", spec["label"])),
                            "process_id": 1234, "process_group_id": 1234}
                matrix._capture_header(spec, capture, started, terminal, build)
                terminal["argv"] = ["tampered"]
                with self.assertRaises(matrix.ColdMatrixError):
                    matrix._capture_header(spec, capture, started, terminal, build)

    def test_protocol_is_immutable_and_byte_bound(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            with _frozen_protocol(Path(directory)) as path:
                self.assertEqual(
                    path.read_bytes(),
                    matrix._canonical_bytes(matrix.protocol_value()) + b"\n",
                )
                path.write_bytes(path.read_bytes().replace(b'"cpu":2', b'"cpu":3'))
                with self.assertRaises(matrix.ColdMatrixError):
                    matrix._protocol_binding()

    def test_elapsed_tail_tampering_is_rejected(self) -> None:
        raw = list(range(1, 31))
        stats = _stats(raw)
        stats["p99"] = stats["p99"] - 1
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._check_stats(stats, raw, "stats")

    def test_cold_proof_requires_positive_read_bytes(self) -> None:
        proof = _eligible_proof()
        matrix._check_cold_sample(proof, "proof")
        proof["read_bytes_delta"] = 0
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._check_cold_sample(proof, "proof")

    def test_cold_proof_requires_zero_residency(self) -> None:
        proof = _eligible_proof()
        proof["dirty_bytes"] = 4096
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._check_cold_sample(proof, "proof")

    def test_timed_text_oracle_cannot_fall_back_to_untimed_replay(self) -> None:
        sample = _sample(0)
        sample.pop("docx_timed_full_text_sha256")
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._text_observation(sample, "sample")

    def test_report_rejects_missing_timed_text_oracle(self) -> None:
        value = _report()
        value["filesystem_evidence"][0]["samples"][0].pop("docx_timed_full_text_sha256")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ColdMatrixError):
                matrix.validate_report(path, role="normal", cache_state="warm", control=False)

    def test_report_rejects_wrong_cold_advice(self) -> None:
        value = _report("cold-requested")
        value["filesystem_evidence"][0]["samples"][4]["cold_advice"] = "not_requested"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ColdMatrixError):
                matrix.validate_report(path, role="normal", cache_state="cold-requested", control=False)

    def test_prepared_control_must_emit_explicit_ineligibility(self) -> None:
        value = _report()
        value["configuration"]["cases"] = [matrix.CONTROL_CASE]
        value["configuration"]["filesystem_cache_states"] = ["cold-verified"]
        value["results"] = []
        value["filesystem_evidence"][0].update({
            "case": matrix.CONTROL_CASE,
            "cache_states": ["cold-verified"],
            "samples": [],
            "cold_verified_status": "ineligible_prepared_query_control",
        })
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "control.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", cache_state="cold-verified", control=True)
            value["filesystem_evidence"][0]["cold_verified_status"] = "eligible"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ColdMatrixError):
                matrix.validate_report(path, role="normal", cache_state="cold-verified", control=True)

    def test_analysis_has_raw_vectors_and_no_comparison_claim(self) -> None:
        report = _report()
        entry = {
            "spec": {
                "label": "r1-normal-warm",
                "kind": "formal",
                "role": "normal",
                "repeat": 1,
                "cache_state": "warm",
                "case": matrix.CASE,
                "samples": 30,
                "warmups": 3,
            },
            "terminal_sha256": "a" * 64,
            "report_sha256": "b" * 64,
            "report": report,
            "resource": {},
        }
        builds = {
            role: {
                "receipt_sha256": f"{index}" * 64,
                "binary": {},
                "source": {},
            }
            for index, role in enumerate(matrix.ROLES, 1)
        }
        summary = matrix.analyze_data([entry], builds)
        self.assertIn("raw_vectors", summary["rows"][0])
        self.assertIn("percentiles", summary["rows"][0])
        self.assertIn("adverse_rows", summary)
        self.assertFalse(summary["claim_authorized"])
        self.assertEqual(summary["performance_claim"], "none")
        self.assertNotIn("before", json.dumps(summary).lower())
        self.assertNotIn("after", json.dumps(summary).lower())


class ReplayProofTests(unittest.TestCase):
    def test_raw_overlap_and_query_io_cannot_hide_behind_classification(self):
        matrix._check_replay(_replay(), "test")
        for field, value in [("query_read_calls", 1), ("open_media_payload_overlap_bytes", 1),
                             ("materializations", 1), ("semantic_sha256", "0" * 64)]:
            altered = _replay()
            altered[field] = value
            with self.subTest(field=field), self.assertRaises(matrix.ColdMatrixError):
                matrix._check_replay(altered, "test")

    def test_aligned_probe_retains_exact_raw_overlap(self):
        replay = _replay()
        probe = dict(matrix.EXPECTED_ALIGNED_TAIL_PROBE)
        replay.update(source_bytes=probe["aligned_source_bytes"],
                      source_sha256=probe["aligned_source_sha256"],
                      classification=matrix.EXPECTED_ALIGNED_REPLAY_CLASSIFICATION,
                      aligned_eocd_tail_probe=probe)
        for category in ("main", "media", "unselected", "core"):
            for suffix in ("overlap_bytes", "covered_bytes"):
                replay[f"open_{category}_payload_{suffix}"] = probe[f"eocd_tail_probe_{category}_payload_overlap_bytes"]
        kwargs = dict(source_bytes=probe["aligned_source_bytes"], source_sha256=probe["aligned_source_sha256"], aligned_tail_probe=True)
        matrix._check_replay(replay, "test", **kwargs)
        replay["open_media_payload_overlap_bytes"] = 0
        with self.assertRaises(matrix.ColdMatrixError):
            matrix._check_replay(replay, "test", **kwargs)


if __name__ == "__main__":
    unittest.main()
