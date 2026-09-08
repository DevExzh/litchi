#!/usr/bin/env python3
"""Focused unit tests for the 0469 analysis and verification helpers."""

from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path
import tempfile
import unittest

import analyze
import verify


class PathAndJsonTests(unittest.TestCase):
    def test_traversal_and_nonfinite_json_are_rejected(self) -> None:
        with self.assertRaises(verify.VerificationError):
            verify.relative("../outside", "path")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "nonfinite.json"
            path.write_text('{"value": NaN}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.read_json(path, "nonfinite.json")

    def test_sha256_file_is_content_addressed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "payload.bin"
            payload = b"0469 evidence"
            path.write_bytes(payload)
            self.assertEqual(
                verify.sha256_file(path),
                (hashlib.sha256(payload).hexdigest(), len(payload)),
            )


class ResourceParserTests(unittest.TestCase):
    def test_gnu_time_rss_extracts_kib_and_excludes_latency(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "resource.log"
            path.write_text(
                "Maximum resident set size (kbytes): 104,388\n",
                encoding="utf-8",
            )
            parsed = analyze.parse_gnu_rss(path, root=Path(directory))
        self.assertEqual(parsed["maximum_resident_set_kib"], 104388)
        self.assertEqual(parsed["maximum_resident_set_bytes"], 104388 * 1024)
        self.assertEqual(parsed["latency_comparison"], "excluded")

    def test_heaptrack_parser_keeps_rounded_peak_as_display_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "heaptrack-print.stdout"
            path.write_text(
                "calls to allocation functions: 12 (3/s)\n"
                "temporary memory allocations: 57 (4/s)\n"
                "peak heap memory consumption: 104.38M\n"
                "peak RSS (including heaptrack overhead): 120.00M\n",
                encoding="utf-8",
            )
            parsed = analyze.parse_heaptrack(path, root=Path(directory))
        self.assertEqual(parsed["allocation_calls"], 12)
        self.assertEqual(parsed["temporary_allocations"], 57)
        self.assertEqual(parsed["peak_heap_display"], "104.38M")
        self.assertNotIn("peak_rss_bytes", parsed)
        self.assertEqual(parsed["latency_comparison"], "excluded")

    def test_heaptrack_fractional_count_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "heaptrack-print.stdout"
            path.write_text(
                "calls to allocation functions: 12.5\n"
                "temporary memory allocations: 57\n"
                "peak heap memory consumption: 1.00M\n",
                encoding="utf-8",
            )
            with self.assertRaises(analyze.AnalysisError):
                analyze.parse_heaptrack(path, root=Path(directory))


class EvidencePolicyTests(unittest.TestCase):
    def test_seal_requires_exact_coverage_and_authentic_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            member = root / "member.json"
            member.write_text("{}\n", encoding="utf-8")
            checksum = hashlib.sha256(member.read_bytes()).hexdigest()
            (root / "SHA256SUMS").write_text(
                f"{checksum}  member.json\n", encoding="utf-8"
            )
            self.assertEqual(verify.verify_sha256sums(root), 1)
            member.write_text('{"tampered":true}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_sha256sums(root)
            member.write_text("{}\n", encoding="utf-8")
            (root / "unlisted.txt").write_text("extra\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.verify_sha256sums(root)

    def test_chronology_rejects_overlapping_capture_intervals(self) -> None:
        def receipt(start: str, finish: str) -> dict[str, str]:
            return {"started_utc": start, "finished_utc": finish}

        ordered = {
            lane: receipt(
                f"2026-09-08T00:00:{index * 2:02d}+00:00",
                f"2026-09-08T00:00:{index * 2 + 1:02d}+00:00",
            )
            for index, lane in enumerate(verify.LANES)
        }
        exports = {
            "A-heap": receipt("2026-09-08T00:01:00+00:00", "2026-09-08T00:01:01+00:00"),
            "B-heap": receipt("2026-09-08T00:01:02+00:00", "2026-09-08T00:01:03+00:00"),
        }
        ordered["B1"]["started_utc"] = "2026-09-08T00:00:00.5+00:00"
        with self.assertRaises(verify.VerificationError):
            verify.verify_chronology(Path("/tmp"), ordered, None, exports)

    def test_candidate_build_can_follow_first_control_capture(self):
        ordered = {
            lane: {'started_utc': f'2026-09-08T00:{i:02d}:00+00:00',
                   'finished_utc': f'2026-09-08T00:{i:02d}:30+00:00'}
            for i, lane in enumerate(verify.LANES)
        }
        exports = {
            lane: {'started_utc': f'2026-09-08T00:{i:02d}:00+00:00',
                   'finished_utc': f'2026-09-08T00:{i:02d}:30+00:00'}
            for i, lane in enumerate(verify.HEAP_LANES, 8)
        }
        build = {'started_utc': '2026-09-08T00:00:31+00:00',
                 'finished_utc': '2026-09-08T00:00:59+00:00'}
        result = verify.verify_chronology(Path('/tmp'), ordered, build, exports)
        self.assertTrue(result['build_before_candidate_captures'])
        build['finished_utc'] = '2026-09-08T00:01:01+00:00'
        with self.assertRaises(verify.VerificationError):
            verify.verify_chronology(Path('/tmp'), ordered, build, exports)

    def test_canonical_json_rejects_nonfinite_values(self) -> None:
        with self.assertRaises(analyze.AnalysisError):
            analyze.canonical({"value": float("nan")})


class RegressionTests(unittest.TestCase):
    def test_source_projection_preserves_raw_and_measured_vectors(self):
        report = {'results': [{'source': {'read_calls': [], 'read_bytes': [7]},
                               'elapsed_ns': {'samples': [1, 2, 3]},
                               'operation_metrics': {'values': []}}]}
        projected, removed = analyze.comparison_projection(report)
        self.assertEqual(removed, ['results/0/source/read_calls'])
        self.assertEqual(report['results'][0]['source']['read_calls'], [])
        self.assertNotIn('read_calls', projected['results'][0]['source'])
        self.assertEqual(projected['results'][0]['source']['read_bytes'], [7])
        self.assertEqual(projected['results'][0]['operation_metrics'], {'values': []})
        self.assertEqual(projected['results'][0]['elapsed_ns'], {'samples': [1, 2, 3]})

    def test_reused_control_must_match_prior_binary(self):
        root = Path(__file__).resolve().parent
        with tempfile.TemporaryDirectory() as directory:
            temporary = Path(directory)
            bundle = temporary / 'change-0469'
            prior = temporary / 'change-0468'
            (bundle / 'sources').mkdir(parents=True)
            prior.mkdir()
            for name in ('binding.json', 'build.json', 'source-binding.json'):
                shutil.copy2(root.parent / 'change-0468' / name, prior / name)
            shutil.copy2(root / 'sources/control.json', bundle / 'sources/control.json')
            shutil.copy2(root / 'control-binding.json', bundle / 'control-binding.json')
            binding = verify.verify_role_binding(bundle, 'control')
            binding['bytes'] += 1
            (bundle / 'control-binding.json').write_text(json.dumps(binding))
            with self.assertRaisesRegex(verify.VerificationError, 'prior binding: bytes differs'):
                verify.verify_role_binding(bundle, 'control')


if __name__ == "__main__":
    raise SystemExit(unittest.main())
