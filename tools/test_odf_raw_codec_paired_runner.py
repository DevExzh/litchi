#!/usr/bin/env python3
"""Source-only tests for the ODF raw-codec manifest and paired runner."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import re
import tempfile
import unittest

from tools.native_qualification import odf_raw_codec_paired_runner as runner


ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = ROOT / "tools/native_qualification/manifests/odf-raw-codec.json"


def manifest() -> dict:
    return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))


class OdfRawCodecRunnerTests(unittest.TestCase):
    def test_manifest_binds_all_routes_and_exact_protocol(self) -> None:
        value = runner.validate_manifest(manifest())
        self.assertEqual([route["id"] for route in value["routes"]], list(runner.ROUTE_IDS))
        self.assertEqual(value["paired"]["abba_order"], list(runner.ABBA_ORDER))
        self.assertEqual(value["paired"]["warmup_iterations_per_leg"], runner.WARMUPS)
        self.assertEqual(value["paired"]["samples_per_leg"], runner.SAMPLES)
        self.assertEqual(value["paired"]["probe_iterations"], runner.PROBE_ITERATIONS)
        self.assertEqual(value["source"]["candidate_overlay"], [])
        self.assertEqual(value["source"]["candidate_mode"], runner.CANDIDATE_MODE)
        self.assertEqual(value["source"]["candidate_integrated_files"], [str(runner.CONSUMER_MANIFEST), str(runner.LOCK_PATH), str(runner.PROBE_PATH)])
        self.assertEqual(value["source"]["protocol_p_commit"], runner.PROTOCOL_P_COMMIT)
        self.assertEqual(value["source"]["protocol_p_tree"], runner.PROTOCOL_P_TREE)
        self.assertIn(runner.NON_ACCEPTANCE, value["claim_scope"])

    def test_manifest_hashes_match_live_adapter_sources(self) -> None:
        value = manifest()

        def digest(relative: str) -> str:
            return hashlib.sha256((ROOT / relative).read_bytes()).hexdigest()

        self.assertEqual(value["probe"]["sha256"], digest(str(runner.PROBE_PATH)))
        self.assertEqual(
            value["source"]["candidate_consumer_manifest_sha256"],
            digest(str(runner.CONSUMER_MANIFEST)),
        )
        self.assertEqual(value["source"]["candidate_consumer_manifest_bytes"], (ROOT / runner.CONSUMER_MANIFEST).stat().st_size)
        self.assertEqual(value["source"]["candidate_lock_sha256"], digest(str(runner.LOCK_PATH)))
        self.assertEqual(value["source"]["candidate_lock_bytes"], (ROOT / runner.LOCK_PATH).stat().st_size)
        self.assertEqual(value["probe"]["bytes"], (ROOT / runner.PROBE_PATH).stat().st_size)
        self.assertEqual(value["build"], runner.EXPECTED_BUILD)

    def test_rust_route_descriptors_match_manifest(self) -> None:
        source = (ROOT / "tools/perf-baseline/examples/odf_raw_codec_qualification.rs").read_text(encoding="utf-8")
        descriptor_source = source[source.index("fn descriptor"):source.index("fn proof")]
        pattern = re.compile(
            r'\s*"(?P<id>[^"]+)" => json!\(\{\s*'
            r'"id": id,\s*'
            r'"family": "(?P<family>[^"]+)",\s*'
            r'"operations": \[(?P<operations>.*?)\],\s*'
            r'"boundary": "(?P<boundary>[^"]+)",\s*'
            r'\}\),',
            re.DOTALL,
        )
        rust_routes = []
        for match in pattern.finditer(descriptor_source):
            rust_routes.append(
                {
                    "id": match.group("id"),
                    "family": match.group("family"),
                    "operations": re.findall(r'"([^"]+)"', match.group("operations")),
                    "boundary": match.group("boundary"),
                }
            )
        self.assertEqual(rust_routes, manifest()["routes"])

    def test_manifest_rejects_stale_hashes_and_route_mutation(self) -> None:
        value = manifest()
        value["source"]["candidate_lock_sha256"] = None
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)
        value = manifest()
        value["source"]["candidate_lock_bytes"] += 1
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)
        value = manifest()
        value["build"]["profile"] = "dev"
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)
        value = manifest()
        value["routes"][0], value["routes"][1] = value["routes"][1], value["routes"][0]
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)
        value = manifest()
        value["candidate_commit"] = "a" * 40
        with self.assertRaises(runner.RunnerError):
            runner.validate_manifest(value)

    def test_manifest_rejects_boundary_setup_and_proof_mutations(self) -> None:
        for path, mutation in (
            (("routes", 0, "operations"), lambda value: value.append("unreviewed")),
            (("routes", 1, "boundary"), lambda value: value + " changed"),
            (("phases", "setup_outside_clock"), lambda value: value[:-1]),
            (("phases", "output_proof"), lambda value: value[:-1]),
        ):
            value = manifest()
            target = value
            for key in path[:-1]:
                target = target[key]
            key = path[-1]
            changed = mutation(target[key])
            if changed is not None:
                target[key] = changed
            with self.assertRaises(runner.RunnerError, msg=f"mutation {path} was accepted"):
                runner.validate_manifest(value)

    def test_build_journal_is_variant_major_and_hashed(self) -> None:
        events = [{"event_id": f"build-{key.replace('/', '-')}"} for key in runner._build_keys()]
        journal = runner._journal(events)
        self.assertEqual(
            [event["event_id"] for event in journal["events"]],
            ["build-baseline-normal", "build-candidate-normal", "build-baseline-allocator", "build-candidate-allocator"],
        )
        self.assertEqual(runner._journal_events(journal)["build-candidate-allocator"], events[-1])
        broken = copy.deepcopy(journal)
        broken["events"][1], broken["events"][2] = broken["events"][2], broken["events"][1]
        with self.assertRaises(runner.RunnerError):
            runner._journal_events(broken)

    def test_ordinary_and_allocator_commands_use_the_optimized_profile(self) -> None:
        instance = runner.PairedRunner.__new__(runner.PairedRunner)
        instance.target_root = Path("/tmp/odf-raw-release-target")
        normal = instance._build_command("baseline", "normal")
        allocator = instance._build_command("candidate", "allocator")
        self.assertIn("--release", normal)
        self.assertIn("--release", allocator)
        self.assertNotIn("--debug", normal)
        self.assertEqual(
            instance._binary("baseline", "normal"),
            instance._target("baseline", "normal") / "release" / "examples" / runner.PROBE_NAME,
        )
        self.assertEqual(
            instance._binary("candidate", "allocator"),
            instance._target("candidate", "allocator") / "release" / "examples" / runner.PROBE_NAME,
        )
        self.assertIn("raw-codec-allocator", allocator[-1])

    def test_fresh_file_rejects_existing_and_symlinked_ancestors(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-raw-runner-") as directory:
            root = Path(directory)
            existing = root / "existing"
            existing.write_text("already", encoding="utf-8")
            instance = runner.PairedRunner.__new__(runner.PairedRunner)
            instance.baseline_root = root / "baseline"
            instance.candidate_root = root / "candidate"
            with self.assertRaises(runner.RunnerError):
                instance._fresh_file(existing, "existing")
            target = root / "target"
            target.mkdir()
            link = root / "link"
            link.symlink_to(target, target_is_directory=True)
            with self.assertRaises(runner.RunnerError):
                instance._fresh_file(link / "result.json", "symlinked output")

    def test_probe_source_contains_all_negative_and_fidelity_proofs(self) -> None:
        probe = (ROOT / "tools/perf-baseline/examples/odf_raw_codec_qualification.rs").read_text(encoding="utf-8")
        for route in runner.ROUTE_IDS:
            self.assertIn(route, probe)
        for case in manifest()["phases"]["negative_cases"]:
            self.assertIn(case, probe)
        for proof in ("source_bytes_ok", "facade_bytes_ok", "untouched_members_ok", "negative_cases_ok"):
            self.assertIn(f'"{proof}"', probe)
        for fidelity_field in (
            "archive_comment_sha256",
            "archive_comment_bytes",
            "name_raw_sha256",
            "content_sha256",
            "content_bytes",
            "declared_uncompressed_bytes",
            "compressed_bytes",
            "crc32",
            "version_needed",
            "flags",
            "external_attributes",
            "comment_raw_sha256",
            "comment_bytes",
        ):
            self.assertIn(fidelity_field, probe)
        self.assertIn("select exactly one raw-codec API era", probe)
        self.assertIn("NAMESPACE_CONTRACT", probe)
        self.assertIn("legacy_unknown_namespace_event", probe)
        self.assertIn("document_rejects_undeclared_namespace", probe)
        self.assertNotIn("Cursor::new(opaque.to_vec())", probe)
        self.assertNotIn("timed_fragment.clone()", probe)
        self.assertIn("one owned", probe)
        self.assertIn("let mut timed_inputs: Vec<Vec<Vec<u8>>> = Vec::with_capacity(iterations);", probe)
        self.assertIn("open_bytes_with(bytes,", probe)
        self.assertNotIn("let file = open_bytes_with(bytes.clone(), &OpenOptions::new())", probe)
        self.assertIn("PackageCompression::Stored", probe)

    def test_probe_report_requires_real_row_and_measurement_schema_without_following_symlinks(self) -> None:
        instance = runner.PairedRunner.__new__(runner.PairedRunner)
        instance.manifest = runner.validate_manifest(manifest())
        proof = {key: True for key in (
            "operation_ok", "changed_ok", "no_op_ok", "reopen_ok", "source_bytes_ok",
            "facade_bytes_ok", "untouched_members_ok", "manifest_ok", "metadata_ok",
            "namespace_ok", "bounded_error_ok", "negative_cases_ok",
        )}
        negative = {key: True for key in instance.manifest["phases"]["negative_cases"]}
        routes = []
        for descriptor in instance.manifest["routes"]:
            routes.append({
                "route": descriptor["id"],
                "family": descriptor["family"],
                "descriptor": descriptor,
                "mode": "timed",
                "iterations": 2,
                "details": {"proof": proof},
                "timing": {
                    "status": "measured",
                    "elapsed_claimable": True,
                    "allocator_enabled": False,
                    "measurement": {"iterations": 2, "ns_per_operation": 1, "allocation_samples": []},
                },
            })
        payload = {
            "schema": "litchi.odf.raw-codec-qualification.v1",
            "mode": "timed",
            "iterations": 2,
            "routes": routes,
            "families": list(runner.FAMILY_IDS),
            "negative_cases": negative,
            "api_era": "current",
            "namespace_contract": runner.NAMESPACE_CONTRACTS["candidate"],
            "claim_scope": runner.NON_ACCEPTANCE,
            "allocator_enabled": False,
        }
        with tempfile.TemporaryDirectory(prefix="odf-raw-report-") as directory:
            root = Path(directory)
            report = root / "report.json"
            report.write_text(json.dumps(payload) + "\n", encoding="utf-8")
            self.assertEqual(instance._validate_probe_output(report, "timed", 2, "candidate")["iterations"], 2)
            minimal = copy.deepcopy(payload)
            del minimal["routes"][0]["family"]
            minimal["routes"][0]["timing"]["measurement"]
            report.write_text(json.dumps(minimal) + "\n", encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                instance._validate_probe_output(report, "timed", 2, "candidate")
            target = root / "target.json"
            target.write_text(json.dumps(payload) + "\n", encoding="utf-8")
            report.unlink()
            report.symlink_to(target)
            with self.assertRaises(runner.RunnerError):
                instance._validate_probe_output(report, "timed", 2, "candidate")


if __name__ == "__main__":
    unittest.main()
