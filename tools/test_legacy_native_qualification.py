#!/usr/bin/env python3
"""Build-free tests for the legacy native qualification protocol."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import native_qualification_runner as runner  # noqa: E402
from native_qualification import legacy_native_validator as validator  # noqa: E402
from native_qualification import prepare_legacy_native_inputs as preparation  # noqa: E402


def _record(path: Path) -> dict[str, object]:
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


class LegacyNativeQualificationTests(unittest.TestCase):
    def test_reviewed_rtf_probe_pin_matches_tracked_source(self) -> None:
        probe = ROOT / validator.REVIEWED_INPUTS["legacy-rtf"]["probe_relative"]
        digest = hashlib.sha256(probe.read_bytes()).hexdigest()
        self.assertEqual(digest, validator.REVIEWED_INPUTS["legacy-rtf"]["probe_sha256"]["after"])
        self.assertEqual(
            preparation.LANES["legacy-rtf"]["probe_sha256"]["after"],
            digest,
        )

    def test_reviewed_ole_adapter_materializes_the_candidate_probe_hash(self) -> None:
        source = subprocess.check_output(
            [
                "git",
                "show",
                "fdb139197d08677d6fa0e14d536e798e380e4cb1:tools/perf-baseline/examples/ole_w5_probe.rs",
            ],
            cwd=ROOT,
        ).decode("utf-8")
        baseline = preparation.materialize_ole_baseline_adapter(source)
        adapted = preparation.materialize_ole_adapter(source)
        baseline_recipe = preparation.OLE_BASELINE_ADAPTER_RECIPE
        candidate_recipe = preparation.OLE_ADAPTER_RECIPE
        self.assertEqual(hashlib.sha256(source.encode()).hexdigest(), baseline_recipe["raw_sha256"])
        self.assertEqual(hashlib.sha256(baseline.encode()).hexdigest(), baseline_recipe["adapted_sha256"])
        self.assertEqual(hashlib.sha256(adapted.encode()).hexdigest(), candidate_recipe["after_sha256"])
        self.assertEqual(
            validator.REVIEWED_INPUTS["legacy-ole-properties"]["probe_sha256"],
            {"before": baseline_recipe["adapted_sha256"], "after": candidate_recipe["after_sha256"]},
        )
        self.assertEqual(
            preparation.OLE_ADAPTER_RECORD,
            validator.REVIEWED_INPUTS["legacy-ole-properties"]["adapter"],
        )
        self.assertEqual(candidate_recipe["before_sha256"], hashlib.sha256(baseline.encode()).hexdigest())
        self.assertIn('current-owner-api', baseline)
        self.assertIn('current-property-api', baseline)
        self.assertIn('legacy-property-api', baseline)
        self.assertEqual(baseline.count('ObjectPackage::open_with_options('), 1)
        self.assertEqual(baseline.count('Ok(source.edit()?)'), 1)
        self.assertEqual(baseline.count('package.write_to(&mut output)?'), 1)
        self.assertEqual(baseline.count('UserProperties::with_limits'), 1)
        self.assertEqual(baseline.count('UserEdit::with_limits'), 1)

    def test_manifests_bind_reusable_drivers_and_contracts(self) -> None:
        expected = {
            "legacy-rtf.json": ("legacy-rtf", 5, ["routes"]),
            "legacy-ole-properties.json": ("legacy-ole-properties", 1, ["read", "edit"]),
            "legacy-xls-properties.json": ("legacy-xls-properties", 1, ["metadata"]),
        }
        for name, (lane, blocks, programs) in expected.items():
            self.assertNotRegex(
                (ROOT / "tools/native_qualification/manifests" / name).read_text(encoding="utf-8"),
                r"(?<![A-Za-z0-9])[0-9a-f]{40}(?![A-Za-z0-9])",
            )
            data = json.loads(
                (ROOT / "tools/native_qualification/manifests" / name).read_text(encoding="utf-8")
            )
            self.assertEqual(data["lane"], lane)
            self.assertIn("preparation_receipt", data)
            self.assertEqual(data.get("abba_blocks", 1), blocks)
            self.assertEqual([item["id"] for item in data["legs"]["before"]["programs"]], programs)
            self.assertEqual([item["id"] for item in data["legs"]["after"]["programs"]], programs)
            for leg in ("before", "after"):
                self.assertTrue(data["legs"][leg]["allow_dirty"])
                self.assertTrue(data["legs"][leg]["allowed_dirty_paths"])

    def test_validator_rejects_fabricated_five_block_receipt_without_provenance(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-rft-") as directory:
            root = Path(directory)
            output = root / "output"
            output.mkdir()
            targets = {"before": root / "targets" / "before", "after": root / "targets" / "after"}
            source_heads = {
                "before": validator.CONTRACTS["legacy-rtf"]["baseline"],
                "after": "1" * 40,
            }
            identities: dict[str, object] = {}
            for leg, head in source_heads.items():
                source = root / f"source-{leg}"
                lock = source / "Cargo.lock"
                probe = source / "crates/litchi-rtf/examples/refusal_contract_bench.rs"
                probe.parent.mkdir(parents=True)
                lock.write_text(f"lock-{leg}\n", encoding="utf-8")
                probe.write_text(f"probe-{leg}\n", encoding="utf-8")
                target = targets[leg]
                target.mkdir(parents=True)
                binary = target / "release/examples/refusal_contract_bench"
                binary.parent.mkdir(parents=True)
                binary.write_bytes(f"binary-{leg}\n".encode())
                binary.chmod(0o755)
                files = [_record(lock), _record(probe)]
                identities[leg] = {
                    "source_root": str(source),
                    "head": head,
                    "tree": "a" * 40,
                    "target_dir": str(target),
                    "output_dir": str(output),
                    "identity_sha256": "b" * 64,
                    "lock_paths": [str(lock)],
                    "probe_paths": [str(probe)],
                    "identity_files": files,
                    "artifacts": [_record(binary)],
                }
            receipt = {
                "schema": validator.SCHEMA,
                "status": "passed",
                "phase": "build",
                "lane": "legacy-rtf",
                "abba_order": list(validator.ABBA_ORDER),
                "abba_blocks": 5,
                "limits": {
                    "target_preflight_bytes": validator.TARGET_PRE_FLIGHT,
                    "target_hard_bytes": validator.TARGET_HARD,
                    "memory_bytes": validator.MEMORY_HARD,
                },
                "identities": identities,
            }
            (output / "build-receipt.json").write_text(json.dumps(receipt), encoding="utf-8")
            events = []
            routes = validator.CONTRACTS["legacy-rtf"]["routes"]
            for block in range(1, 6):
                for index, leg in enumerate(validator.ABBA_ORDER, start=1):
                    stdout = output / f"rtf-routes-block-{block}-{leg}-{index}.jsonl"
                    stdout.write_text(
                        "".join(json.dumps({"scenario": route, "median_us": 1.0}) + "\n" for route in routes),
                        encoding="utf-8",
                    )
                    stderr = output / f"stderr-{block}-{leg}-{index}"
                    stderr.write_bytes(b"")
                    events.append(
                        {
                            "schema": validator.SCHEMA,
                            "event_id": f"abba-block-{block:02d}-{index:02d}-{leg}-routes",
                            "leg": leg,
                            "source_head": source_heads[leg],
                            "returncode": 0,
                            "runner_error": None,
                            "process_guard_violations": [],
                            "post_exit_descendants": [],
                            "post_exit_offenders": [],
                            "environment": {
                                "CARGO_TARGET_DIR": str(targets[leg]),
                                "CARGO_BUILD_JOBS": "1",
                                "CARGO_INCREMENTAL": "0",
                                "LITCHI_MEM_MAX": "20G",
                            },
                            "target": {"path": str(targets[leg]), "bytes_after": 0},
                            "stdout": _record(stdout),
                            "stderr": _record(stderr),
                        }
                    )
            (output / "events-abba.jsonl").write_text(
                "".join(json.dumps(event) + "\n" for event in events), encoding="utf-8"
            )
            with self.assertRaisesRegex(validator.ValidationError, "manifest path is not source-bound"):
                validator.validate_directory(
                    output,
                    "legacy-rtf",
                    source_heads["before"],
                    source_heads["after"],
                    targets["before"],
                    targets["after"],
                )

    def test_text_executable_is_not_a_native_artifact(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-text-artifact-") as directory:
            path = Path(directory) / "probe"
            path.write_text("#!/bin/sh\nexit 0\n", encoding="utf-8")
            path.chmod(0o755)
            with self.assertRaisesRegex(validator.ValidationError, "ELF executable"):
                validator._validate_hash_record(_record(path), "artifact", require_executable=True)

    def test_symlinked_artifact_path_is_rejected_before_resolution(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-symlink-artifact-") as directory:
            root = Path(directory)
            target = root / "real"
            target.write_bytes(b"\x7fELF qualification fixture")
            target.chmod(0o755)
            link = root / "link"
            link.symlink_to(target)
            with self.assertRaisesRegex(validator.ValidationError, "contains a symlink"):
                validator._validate_hash_record(_record(link), "artifact", require_executable=True)

    def test_validator_rejects_missing_rtf_route(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-rft-invalid-") as directory:
            # The valid-fixture setup is intentionally kept in the preceding
            # test; this direct parser check proves the route set is fail-closed
            # without running a native process.
            path = Path(directory) / "routes.jsonl"
            path.write_text(json.dumps({"scenario": "open_1k", "median_us": 1.0}) + "\n", encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "workload identity"):
                validator._validate_rtf_log(path, validator.CONTRACTS["legacy-rtf"]["routes"])

    def test_rtf_parser_requires_exact_refusal_kind_reason_and_fidelity(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-rft-contract-") as directory:
            path = Path(directory) / "routes.jsonl"
            rows = []
            for scenario in validator.CONTRACTS["legacy-rtf"]["routes"]:
                row = {
                    "scenario": scenario,
                    "median_us": 1.0,
                    "workload": "rtf-refusal-contract-v1",
                    **validator.RTF_ROUTE_CONTRACT[scenario],
                }
                rows.append(row)
            path.write_text("".join(json.dumps(row) + "\n" for row in rows), encoding="utf-8")
            report = validator._validate_rtf_log(path, validator.CONTRACTS["legacy-rtf"]["routes"])
            self.assertEqual(report["fixture_id"], "rtf-refusal-contract-v1")
            rows[-1]["reason"] = "different"
            path.write_text("".join(json.dumps(row) + "\n" for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "fixed refusal contract"):
                validator._validate_rtf_log(path, validator.CONTRACTS["legacy-rtf"]["routes"])

    def test_csv_parser_requires_all_retained_samples(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-csv-") as directory:
            path = Path(directory) / "read.csv"
            path.write_text(
                "route,sample,nanoseconds,output_bytes\nread,0,1,10\n",
                encoding="utf-8",
            )
            with self.assertRaisesRegex(validator.ValidationError, "sample count"):
                validator._validate_csv_log(path, "read", 2)

    def test_metadata_parser_requires_doc_xls_ppt_rows(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-metadata-") as directory:
            path = Path(directory) / "metadata.jsonl"
            path.write_text(
                "".join(
                    json.dumps(
                        {
                            "case": case,
                            "bytes": 100,
                            "repeats": 50,
                            "native_us": 1.0,
                            "unified_us": 1.0,
                            "overhead_pct": 0.0,
                        }
                    )
                    + "\n"
                    for case in ("doc", "xls", "ppt")
                ),
                encoding="utf-8",
            )
            report = validator._validate_metadata_log(path, ("doc", "xls", "ppt"))
            self.assertEqual(report["cases"], ["doc", "xls", "ppt"])


if __name__ == "__main__":
    unittest.main()
