#!/usr/bin/env python3
"""Focused unit tests for the 0470 analysis and verification helpers."""

from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path
import tempfile
import unittest

import analyze
import guard_review
import rss_review
import review
import verify


class RelocatedReceiptTests(unittest.TestCase):
    def check_relocation(self, lane, protocol_name, check):
        source = Path(__file__).resolve().parent
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "relocated"
            shutil.copytree(source, root)
            protocol = verify.read_json(root / protocol_name, "protocol")
            bindings = {role: verify.read_json(root / f"{role}-binding.json", role)
                        for role in ("control", "candidate")}
            check(root, lane, protocol, bindings)
            for name in ("started.json", "receipt.json"):
                path = root / lane / name
                item = json.loads(path.read_text())
                item["argv"][item["argv"].index("-o") + 1] = str(root / lane / "resource.log")
                path.write_text(json.dumps(item))
            with self.assertRaisesRegex(verify.VerificationError, "command.*differs"):
                check(root, lane, protocol, bindings)

    def test_guard_relocation_preserves_historical_destination_checks(self):
        self.check_relocation("guard-A1", "guard-protocol.json", guard_review._verify_lane)

    def test_rss_relocation_preserves_historical_destination_checks(self):
        self.check_relocation("rss-A1", "rss-protocol.json",
                              lambda root, lane, protocol, bindings:
                              rss_review._verify_receipt(root, lane, protocol, bindings["control"]))


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
            payload = b"0470 evidence"
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


class RssFollowupTests(unittest.TestCase):
    def test_rss_protocol_binds_its_driver_and_main_protocol(self):
        root = Path(__file__).resolve().parent
        protocol = json.loads((root / 'rss-protocol.json').read_text())
        self.assertEqual(
            protocol['capture_driver_sha256'],
            verify.sha256_file(root / 'rss_capture.py')[0],
        )
        self.assertEqual(
            protocol['main_protocol_sha256'],
            verify.sha256_file(root / 'protocol.json')[0],
        )
        self.assertEqual(protocol['order'], ['rss-A1', 'rss-B1', 'rss-B2', 'rss-A2'])
        self.assertEqual(protocol['samples'], 100)
        self.assertEqual(protocol['warmups'], 5)


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
    def test_historical_review_allows_only_verifier_binding_evolution(self):
        historical = {
            "schema": "litchi-0470-review-v1",
            "full_guard": {"flags_above_threshold": [{"case": "one", "delta_percent": 6.0}]},
            "inputs": {
                "policy": {"sha256": "p"},
                "verifier": {"path": "verify.py", "sha256": "a" * 64, "bytes": 10},
            },
        }
        current = json.loads(json.dumps(historical))
        current["inputs"]["verifier"] = {
            "path": "verify.py",
            "sha256": "b" * 64,
            "bytes": 11,
        }
        evolution = review.historical_metadata_evolution(historical, current)
        self.assertEqual(evolution["allowed_field"], "inputs.verifier")
        self.assertEqual(evolution["historical_verifier"]["sha256"], "a" * 64)
        self.assertEqual(evolution["current_verifier"]["sha256"], "b" * 64)

        current["full_guard"]["flags_above_threshold"][0]["delta_percent"] = 7.0
        with self.assertRaisesRegex(verify.VerificationError, "semantic payload differs"):
            review.historical_metadata_evolution(historical, current)

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
            bundle = temporary / 'change-0470'
            prior = temporary / 'change-0469'
            (bundle / 'sources').mkdir(parents=True)
            prior.mkdir()
            for name in ('candidate-binding.json', 'candidate-source-binding.json', 'build.json'):
                shutil.copy2(root.parent / 'change-0469' / name, prior / name)
            shutil.copy2(root / 'sources/control.json', bundle / 'sources/control.json')
            shutil.copy2(root / 'control-binding.json', bundle / 'control-binding.json')
            binding_path = bundle / 'control-binding.json'
            original_binding_text = binding_path.read_text()
            binding = verify.verify_role_binding(bundle, 'control')
            binding['bytes'] += 1
            binding_path.write_text(json.dumps(binding))
            with self.assertRaisesRegex(verify.VerificationError, 'reused binding: bytes differs'):
                verify.verify_role_binding(bundle, 'control')

            # Refresh the outer reused-binding digest after changing an
            # inner identity field.  Verification must still compare the
            # reused binding's source lineage rather than trusting that outer
            # digest alone.
            binding = json.loads(original_binding_text)
            prior_binding_path = prior / 'candidate-binding.json'
            prior_binding = json.loads(prior_binding_path.read_text())
            prior_binding['source_binding_sha256'] = '0' * 64
            prior_binding_path.write_text(json.dumps(prior_binding))
            binding['reused_binding_sha256'] = hashlib.sha256(
                prior_binding_path.read_bytes()
            ).hexdigest()
            binding_path.write_text(json.dumps(binding))
            with self.assertRaisesRegex(verify.VerificationError, 'reused binding: source_binding_sha256 differs'):
                verify.verify_role_binding(bundle, 'control')

    def test_candidate_manifest_tamper_is_not_hidden_by_refreshed_hashes(self):
        root = Path(__file__).resolve().parent
        with tempfile.TemporaryDirectory() as directory:
            temporary = Path(directory)
            bundle = temporary / 'change-0470'
            prior = temporary / 'change-0469'
            (bundle / 'sources').mkdir(parents=True)
            prior.mkdir()
            for name in ('candidate-binding.json', 'candidate-source-binding.json', 'build.json'):
                shutil.copy2(root.parent / 'change-0469' / name, prior / name)
            shutil.copy2(root / 'control-binding.json', bundle / 'control-binding.json')
            shutil.copy2(root / 'sources/control.json', bundle / 'sources/control.json')

            candidate_manifest = bundle / 'sources/candidate.json'
            control_manifest = bundle / 'sources/control.json'
            candidate_manifest.write_bytes(control_manifest.read_bytes())
            source_binding = json.loads(
                (root.parent / 'change-0469/candidate-source-binding.json').read_text()
            )
            source_binding['schema'] = 'litchi-0470-source-binding-v1'
            source_binding['changed_files'] = []
            source_binding['source_manifest_sha256'] = hashlib.sha256(
                candidate_manifest.read_bytes()
            ).hexdigest()
            candidate_source_path = bundle / 'candidate-source-binding.json'
            candidate_source_path.write_text(json.dumps(source_binding))
            candidate_binding = json.loads((root / 'control-binding.json').read_text())
            candidate_binding.update(
                role='candidate',
                binary_path='/tmp/litchi-goal-0470/candidate',
                source_manifest='sources/candidate.json',
                source_manifest_sha256=source_binding['source_manifest_sha256'],
                source_binding_sha256=hashlib.sha256(
                    candidate_source_path.read_bytes()
                ).hexdigest(),
            )
            for field in ('reused_binding_path', 'reused_binding_sha256', 'prior_build_receipt_path', 'prior_source_binding_path', 'build_receipt_path', 'build_receipt_sha256'):
                candidate_binding.pop(field, None)
            (bundle / 'candidate-binding.json').write_text(json.dumps(candidate_binding))
            bindings = {
                'control': verify.verify_role_binding(bundle, 'control'),
                'candidate': verify.verify_role_binding(bundle, 'candidate'),
            }
            verify.verify_source_manifest_pair(bundle, bindings)

            manifest = json.loads(candidate_manifest.read_text())
            first_name = sorted(manifest)[0]
            manifest[first_name] = '0' * 64
            candidate_manifest.write_text(json.dumps(manifest))
            candidate_binding['source_manifest_sha256'] = hashlib.sha256(
                candidate_manifest.read_bytes()
            ).hexdigest()
            source_binding['source_manifest_sha256'] = candidate_binding['source_manifest_sha256']
            candidate_source_path.write_text(json.dumps(source_binding))
            candidate_binding['source_binding_sha256'] = hashlib.sha256(
                candidate_source_path.read_bytes()
            ).hexdigest()
            (bundle / 'candidate-binding.json').write_text(json.dumps(candidate_binding))
            bindings['candidate'] = verify.verify_role_binding(bundle, 'candidate')
            with self.assertRaisesRegex(verify.VerificationError, 'changed file set differs'):
                verify.verify_source_manifest_pair(bundle, bindings)


if __name__ == "__main__":
    raise SystemExit(unittest.main())
