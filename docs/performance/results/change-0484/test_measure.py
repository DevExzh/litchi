"""Focused acceptance tests for the retained 0484 capture report gate."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import measure  # noqa: E402


def _sha(seed: str) -> str:
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()


def _histogram(*, calls: int = 1, small: int | None = None) -> dict[str, int]:
    return {
        "bytes_0": 0,
        "bytes_1_to_512": calls if small is None else small,
        "bytes_513_to_4096": 0,
        "bytes_4097_to_16384": 0,
        "bytes_16385_to_65536": 0,
        "bytes_over_65536": 0,
    }


def _valid_report(*, role: str = "normal", samples: int = 2) -> tuple[dict, dict, dict]:
    case = measure._case(64, 64, "short", "64")
    source_hash = _sha("a")
    candidate_hash = _sha("b")
    authored_hash = _sha("c")
    source = {
        "archive_bytes": 100,
        "archive_sha256": _sha("d"),
        "main_xml_bytes": 200,
        "main_xml_sha256": source_hash,
        "member_count": 1,
        "semantic": {
            "paragraph_count": 64,
            "order_sha256": _sha("e"),
            "text_sha256": _sha("f"),
            "text_bytes": 64 * 15,
        },
        "members": [{
            "path": "word/document.xml",
            "compression_method": "Deflate",
            "data_descriptor": True,
            "crc32": 1,
            "decoded_bytes": 200,
            "decoded_sha256": source_hash,
            "compressed_bytes": 80,
            "compressed_sha256": _sha("g"),
        }],
        "unchanged_oracle": True,
        "opaque_member_exact": True,
    }
    authored = {
        "authored_count": 64,
        "chunk_mode": "fixed64",
        "text_mode": "short",
        "max_chunk_bytes": 10,
        "replay_window_bytes": 65_536,
        "max_encoded_paragraph_bytes": 20,
        "xml_entity_reference_count": 64,
        "text_bytes": 20 * 64,
        "encoded_xml_bytes": 20 * 64,
        "event_count": 4 * 64,
        "expected_event_sha256": authored_hash,
        "expected_encoded_sha256": _sha("h"),
    }
    oracle = {
        "candidate_archive_bytes": 120,
        "candidate_archive_sha256": candidate_hash,
        "candidate_main_xml_bytes": 240,
        "candidate_main_xml_sha256": candidate_hash,
        "candidate_semantic": {
            "paragraph_count": 128,
            "order_sha256": _sha("i"),
            "text_sha256": _sha("j"),
            "text_bytes": 64 * 15 + 20 * 64,
        },
        "candidate_member_count": 1,
        "candidate_xml_exact": True,
        "candidate_semantic_exact": True,
        "untouched_member_metadata_exact": True,
        "untouched_raw_members_preserved": True,
        "physical_order_exact": True,
        "opaque_member_exact": True,
        "source_unchanged": True,
        "inverse_exact": True,
    }
    limits = {
        "parser_event_limit": 64 * 8 + 128 + 64 * 8 + 64 * 2 + 128,
        "parser_token_bytes": 65_536,
        "parser_workspace_bytes": 13_649_335,
        "max_xml_depth": 16,
        "max_authored_chunk_bytes": 10,
        "replay_window_bytes": 65_536,
    }
    proof = {
        "source_len": 200,
        "source_sha256": source_hash,
        "source_paragraph_count": 64,
        "source_event_count": 300,
        "insertion_offset": 100,
        "candidate_len": 240,
        "candidate_sha256": candidate_hash,
        "candidate_paragraph_count": 128,
        "candidate_event_count": 500,
        "generated_offset": 100,
        "generated_once": True,
        "authored": copy.deepcopy(authored),
    }
    allocation = None
    binary = {
        "binary": "litchi-perf-baseline",
        "allocator": "Rust system allocator",
        "instrumentation": "none",
        "counter_revision": None,
    }
    if role == "allocator":
        binary = {
            "binary": "litchi-perf-baseline-alloc",
            "allocator": "CountingSystemAllocator(std::alloc::System)",
            "instrumentation": "system_allocator_operation_scoped",
            "counter_revision": "serialized_region_peak_v3",
        }
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            "allocation_calls": 10,
            "deallocation_calls": 10,
            "reallocation_calls": 1,
            "failed_allocation_calls": 0,
            "allocated_bytes": 100,
            "deallocated_bytes": 100,
            "live_bytes_before": 500,
            "live_bytes_after": 500,
            "peak_live_bytes_before": 500,
            "peak_live_bytes_after": 700,
            "region_peak_live_bytes": 600,
        }
    sample_list = []
    for index in range(samples):
        sample_list.append({
            "sample": index,
            "source_count": 64,
            "authored_count": 64,
            "chunk_mode": "fixed64",
            "text_mode": "short",
            "elapsed_ns": 100 + index,
            "source_reads": {
                "calls": 1,
                "requested_bytes": 10,
                "returned_bytes": 10,
                "request_histogram": _histogram(),
                "returned_histogram": _histogram(),
            },
            "authored": {
                "opens": 5,
                "events": 4 * 64 * 5,
                "text_chunks": 2 * 64 * 5,
                "text_bytes": 20 * 64 * 5,
            },
            "sink": {
                "accepted_bytes": 120,
                "write_calls": 1,
                "largest_write": 120,
                "histogram": _histogram(),
                "sha256": candidate_hash,
            },
            "allocation": copy.deepcopy(allocation),
            "process": None,
        })
    report = {
        "schema": measure.REPORT_SCHEMA,
        "version": 1,
        "binary": binary,
        "config": {
            "source_counts": [64],
            "authored_counts": [64],
            "chunk_modes": ["fixed64"],
            "text_modes": ["short"],
            "samples": samples,
            "warmups": 1,
            "sink_write_bytes": measure.SINK_WRITE_BYTES,
            "lifecycle": measure.LIFECYCLE,
            "expected_authored_opens": 5,
            "source": measure.SOURCE_CONTRACT,
            "authored_provider": measure.AUTHORED_PROVIDER,
            "sink": measure.SINK_CONTRACT,
            "fixture_dir": None,
        },
        "cases": [{
            "source_count": 64,
            "authored_count": 64,
            "chunk_mode": "fixed64",
            "text_mode": "short",
            "source": source,
            "authored": authored,
            "limits": limits,
            "oracle": oracle,
            "proof": proof,
            "samples": sample_list,
        }],
    }
    return report, case, binary


class ReportAcceptanceTests(unittest.TestCase):
    def _check(
        self,
        report: dict,
        case: dict,
        binary: dict,
        *,
        role: str = "normal",
        samples: int = 2,
        expected_max_replay_bytes: int | None = None,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report_path = root / "report.json"
            report_path.write_text(json.dumps(report), encoding="utf-8")
            runtime_binary = {"path": "/tmp/fake-binary", "sha256": _sha("z"), "bytes": 1}
            argv = measure._argv(runtime_binary, case, samples=samples, warmups=1, report=report_path, resource=root / "resource.txt")
            measure._check_report_shell(
                report_path,
                role,
                case,
                samples=samples,
                warmups=1,
                binary=runtime_binary,
                argv=argv,
                expected_max_replay_bytes=expected_max_replay_bytes,
            )

    def test_valid_normal_report_and_sample_cardinality(self) -> None:
        report, case, binary = _valid_report()
        self._check(report, case, {"path": "/tmp/fake-binary", "sha256": _sha("z"), "bytes": 1})
        report["cases"][0]["samples"].pop()
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, {"path": "/tmp/fake-binary", "sha256": _sha("z"), "bytes": 1})

    def test_shared_shell_requires_one_producer_pass_when_authored_opens_are_zero(self) -> None:
        report, case, _ = _valid_report()
        report["config"]["expected_authored_opens"] = 0
        authored = report["cases"][0]["authored"]
        for sample in report["cases"][0]["samples"]:
            sample["authored"] = {
                "opens": 0,
                "events": authored["event_count"],
                "text_chunks": authored["event_count"] - 2 * authored["authored_count"],
                "text_bytes": authored["text_bytes"],
            }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            report_path = root / "report.json"
            report_path.write_text(json.dumps(report), encoding="utf-8")
            runtime_binary = {"path": "/tmp/fake-binary", "sha256": _sha("z"), "bytes": 1}
            argv = measure._argv(runtime_binary, case, samples=2, warmups=1, report=report_path, resource=root / "resource.txt")
            measure.check_report_shell(
                report_path,
                "normal",
                case,
                samples=2,
                warmups=1,
                binary=runtime_binary,
                argv=argv,
                expected_authored_opens=0,
            )

            report["cases"][0]["samples"][0]["authored"]["events"] = 0
            report_path.write_text(json.dumps(report), encoding="utf-8")
            with self.assertRaises(measure.MeasureError):
                measure.check_report_shell(
                    report_path,
                    "normal",
                    case,
                    samples=2,
                    warmups=1,
                    binary=runtime_binary,
                    argv=argv,
                    expected_authored_opens=0,
                )

    def test_proof_and_limit_identity_are_authenticated(self) -> None:
        report, case, binary = _valid_report()
        report["cases"][0]["proof"]["candidate_sha256"] = _sha("x")
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)

        report, case, binary = _valid_report()
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary, expected_max_replay_bytes=100)
        report, case, binary = _valid_report()
        report["cases"][0]["limits"]["replay_window_bytes"] = 8 * 1024 * 1024
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)

    def test_allocator_live_bytes_must_balance_and_be_measured(self) -> None:
        report, case, binary = _valid_report(role="allocator")
        self._check(report, case, binary, role="allocator")
        report["cases"][0]["samples"][0]["allocation"]["deallocated_bytes"] += 1
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary, role="allocator")

        report, case, binary = _valid_report(role="allocator")
        report["cases"][0]["samples"][0]["allocation"]["failed_allocation_calls"] = 1
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary, role="allocator")

    def test_histogram_bytes_and_largest_write_cannot_be_fabricated(self) -> None:
        report, case, binary = _valid_report()
        sample = report["cases"][0]["samples"][0]
        sample["sink"]["accepted_bytes"] = 600
        report["cases"][0]["oracle"]["candidate_archive_bytes"] = 600
        sample["sink"]["largest_write"] = 100
        sample["sink"]["histogram"] = {
            "bytes_0": 0,
            "bytes_1_to_512": 0,
            "bytes_513_to_4096": 1,
            "bytes_4097_to_16384": 0,
            "bytes_16385_to_65536": 0,
            "bytes_over_65536": 0,
        }
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)

    def test_near_limit_payload_is_exactly_sixty_kibibytes_per_paragraph(self) -> None:
        report, _, binary = _valid_report()
        case = measure._case(64, 64, "near", "64")
        report["config"]["text_modes"] = ["near_limit"]
        observed = report["cases"][0]
        observed["text_mode"] = "near_limit"
        for sample in observed["samples"]:
            sample["text_mode"] = "near_limit"
        authored = observed["authored"]
        authored["text_mode"] = "near_limit"
        authored["text_bytes"] = 64 * measure.TEXT_BYTES["near"] - 64
        authored["max_chunk_bytes"] = 64
        authored["event_count"] = 2 * 64 + ((authored["text_bytes"] // 64 + 63) // 64) * 64
        observed["limits"]["max_authored_chunk_bytes"] = 64
        observed["proof"]["authored"] = copy.deepcopy(authored)
        observed["oracle"]["candidate_semantic"]["text_bytes"] = 64 * 15 + authored["text_bytes"]
        for sample in observed["samples"]:
            sample["authored"]["events"] = authored["event_count"] * 5
            sample["authored"]["text_chunks"] = (authored["event_count"] - 2 * 64) * 5
            sample["authored"]["text_bytes"] = authored["text_bytes"] * 5
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)

    def test_scanner_workspace_bound_is_recomputed_for_target_layout(self) -> None:
        report, case, binary = _valid_report()
        report["cases"][0]["limits"]["parser_workspace_bytes"] += 1
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)

    def test_non_finite_json_numbers_are_rejected(self) -> None:
        report, case, binary = _valid_report()
        report["cases"][0]["samples"][0]["elapsed_ns"] = float("nan")
        with self.assertRaises(measure.MeasureError):
            self._check(report, case, binary)


if __name__ == "__main__":
    unittest.main()
