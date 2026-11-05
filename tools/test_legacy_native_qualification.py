#!/usr/bin/env python3
"""Build-free tests for the legacy native qualification protocol."""

from __future__ import annotations

import hashlib
import json
import os
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

    def test_rtf_adapter_binds_native_variants_and_pinned_locks(self) -> None:
        adapter = preparation.RTF_ADAPTER_RECORD
        reviewed = validator.REVIEWED_INPUTS["legacy-rtf"]
        self.assertEqual(adapter, reviewed["adapter"])
        live = adapter["live_current_probe"]
        live_path = ROOT / live["relative_path"]
        self.assertEqual(hashlib.sha256(live_path.read_bytes()).hexdigest(), live["sha256"])
        self.assertEqual(
            live["lock_sha256"],
            "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
        )
        self.assertEqual(
            hashlib.sha256(Path(live["lock_source"]).read_bytes()).hexdigest(),
            live["lock_sha256"],
        )
        self.assertEqual(live["api"], "current litchi-rtf workspace example; not part of the baseline adapter source")
        self.assertEqual(
            reviewed["probe_sha256"],
            {"before": adapter["probe_sha256"], "after": adapter["probe_sha256"]},
        )
        self.assertEqual(
            reviewed["lock_sha256"],
            {
                "before": "57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f",
                "after": "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c",
            },
        )
        self.assertNotEqual(adapter["probe_sha256"], adapter["historical_failed_compile"]["source_sha256"])
        self.assertNotEqual(live["sha256"], adapter["historical_failed_compile"]["source_sha256"])
        self.assertEqual(adapter["historical_failed_compile"]["exit"], 101)
        self.assertEqual(
            adapter["historical_failed_compile"]["api_context"],
            "current litchi-rtf consumer at candidate fence 35e37dbede6657dd023b114a1536862c18541946",
        )
        self.assertEqual(
            adapter["native_error_api"]["before"],
            "litchi_rtf::edit::Error::{UnsupportedSource,ProtectedDocument}",
        )
        self.assertEqual(
            adapter["native_error_api"]["after"],
            "litchi_rtf::edit::Error::Refused(Refusal)",
        )

    def test_rtf_adapter_rejects_mismatched_feature_variant(self) -> None:
        data = json.loads(
            (ROOT / "tools/native_qualification/manifests/legacy-rtf.json").read_text(encoding="utf-8")
        )
        data["preparation_receipt"] = "/tmp/rtf-preparation.json"
        data["legs"]["before"]["build"][0]["argv"][7] = "rtf-current-api"
        with self.assertRaisesRegex(runner.RunnerError, "reviewed legacy-rtf build route"):
            runner._validate_legacy_native_workload(data)

    def test_rtf_adapter_rejects_stale_probe_identity(self) -> None:
        reviewed = validator.REVIEWED_INPUTS["legacy-rtf"]
        lock = ROOT / reviewed["lock_relative"]
        probe = ROOT / reviewed["probe_relative"]
        identity = {
            "source_root": str(ROOT),
            "lock_paths": [str(lock)],
            "probe_paths": [str(probe)],
            "fixture_paths": [],
            "identity_files": [_record(lock), _record(probe)],
        }
        identity["identity_files"][1]["sha256"] = "0" * 64
        with self.assertRaisesRegex(validator.ValidationError, "SHA-256 changed"):
            validator._validate_reviewed_inputs("legacy-rtf", identity, "after")

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

    def test_ole_adapter_files_reject_symlinks_and_escape_targets(self) -> None:
        source = subprocess.check_output(
            [
                "git",
                "show",
                "fdb139197d08677d6fa0e14d536e798e380e4cb1:tools/perf-baseline/examples/ole_w5_probe.rs",
            ],
            cwd=ROOT,
        ).decode("utf-8")
        with tempfile.TemporaryDirectory(prefix="legacy-native-ole-adapter-") as directory:
            root = Path(directory)
            owned = root / "owned"
            owned.mkdir()
            source_path = root / "source.rs"
            source_path.write_text(source, encoding="utf-8")

            symlink_source = root / "source-link.rs"
            symlink_source.symlink_to(source_path)
            with self.assertRaisesRegex(preparation.PreparationError, "contains a symlink"):
                preparation.materialize_ole_baseline_adapter_file(
                    symlink_source,
                    owned / "symlink-source.rs",
                    owned_root=owned,
                )

            missing = root / "missing.rs"
            dangling_destination = owned / "dangling.rs"
            dangling_destination.symlink_to(missing)
            with self.assertRaisesRegex(preparation.PreparationError, "contains a symlink"):
                preparation.materialize_ole_adapter_file(
                    source_path,
                    dangling_destination,
                    owned_root=owned,
                )
            self.assertFalse(missing.exists())

            parent_link = root / "owned-link"
            parent_link.symlink_to(owned, target_is_directory=True)
            with self.assertRaisesRegex(preparation.PreparationError, "contains a symlink"):
                preparation.materialize_ole_baseline_adapter_file(
                    source_path,
                    parent_link / "parent-link.rs",
                    owned_root=owned,
                )

            outside = root / "outside"
            outside.mkdir()
            with self.assertRaisesRegex(preparation.PreparationError, "escapes owned root"):
                preparation.materialize_ole_adapter_file(
                    source_path,
                    outside / "escape.rs",
                    owned_root=owned,
                )

            baseline_target = owned / "nested" / "baseline.rs"
            preparation.materialize_ole_baseline_adapter_file(source_path, baseline_target, owned_root=owned)
            self.assertEqual(
                hashlib.sha256(baseline_target.read_bytes()).hexdigest(),
                preparation.OLE_BASELINE_ADAPTER_RECIPE["adapted_sha256"],
            )
            candidate_target = owned / "nested" / "candidate.rs"
            preparation.materialize_ole_adapter_file(source_path, candidate_target, owned_root=owned)
            self.assertEqual(
                hashlib.sha256(candidate_target.read_bytes()).hexdigest(),
                preparation.OLE_ADAPTER_RECIPE["after_sha256"],
            )

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

    def test_all_legacy_manifests_load_with_reviewed_workload_features(self) -> None:
        prefixes = {
            "RTF": "legacy-rtf",
            "OLE": "legacy-ole-properties",
            "XLS": "legacy-xls-properties",
        }
        with tempfile.TemporaryDirectory(prefix="legacy-native-manifest-load-") as directory:
            root = Path(directory)
            environment = {
                "LITCHI_COORD_BIN": str(root / "coord-bin"),
                "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
            }
            for prefix in prefixes:
                environment.update(
                    {
                        f"{prefix}_NATIVE_QUAL_OUTPUT": str(root / f"{prefix.lower()}-output"),
                        f"{prefix}_BEFORE_SOURCE": str(root / f"{prefix.lower()}-before"),
                        f"{prefix}_AFTER_SOURCE": str(root / f"{prefix.lower()}-after"),
                        f"{prefix}_TARGET_ROOT": str(root / f"{prefix.lower()}-targets"),
                        f"{prefix}_NATIVE_TMP": str(root / f"{prefix.lower()}-tmp"),
                        f"{prefix}_NATIVE_PREPARATION_RECEIPT": str(root / f"{prefix.lower()}-preparation.json"),
                        f"{prefix}_BASELINE_REVISION": "0" * 40,
                        f"{prefix}_CANDIDATE_REVISION": "1" * 40,
                    }
                )
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                for prefix, lane in prefixes.items():
                    manifest = ROOT / "tools/native_qualification/manifests" / f"{lane}.json"
                    loaded = runner.load_manifest(manifest)
                    self.assertEqual(loaded["lane"], lane)
                    if lane == "legacy-ole-properties":
                        self.assertEqual(
                            loaded["legs"]["before"]["build"][0]["argv"][7],
                            "current-owner-api,legacy-property-api",
                        )
                        self.assertEqual(
                            loaded["legs"]["after"]["build"][0]["argv"][7],
                            "current-owner-api,current-property-api",
                        )
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_preparer_identity_matches_runner_on_real_git_checkout(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-identity-roundtrip-") as directory:
            root = Path(directory) / "source"
            root.mkdir()
            lock = root / "Cargo.lock"
            probe = root / "probe.rs"
            lock.write_text("lock\n", encoding="utf-8")
            probe.write_text("probe\n", encoding="utf-8")
            subprocess.run(["git", "-C", str(root), "init", "-q"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.name", "legacy-test"], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            subprocess.run(["git", "-C", str(root), "commit", "-qm", "identity fixture"], check=True)
            probe_hash = hashlib.sha256(probe.read_bytes()).hexdigest()
            spec = {
                "lock": "Cargo.lock",
                "probe": "probe.rs",
                "fixtures": [],
                "allowed_dirty": [],
                "probe_sha256": {"before": probe_hash, "after": probe_hash},
            }
            prepared = preparation._source_identity(root, "legacy-rtf", "before", spec)
            observed = runner._source_identity(
                {
                    "source_root": str(root),
                    "allow_dirty": False,
                    "lock_paths": ["Cargo.lock"],
                    "probe_paths": ["probe.rs"],
                },
                {"probe_paths": [], "fixture_paths": []},
            )
            for key in (
                "source_root",
                "head",
                "tree",
                "source_tree_sha256",
                "clean",
                "status_paths",
                "status_sha256",
                "diff_sha256",
                "identity_sha256",
                "lock_paths",
                "probe_paths",
                "fixture_paths",
                "identity_files",
            ):
                self.assertEqual(prepared[key], observed[key], key)

    def test_legacy_preparation_receipt_is_external_and_disjoint(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-preparation-paths-") as directory:
            root = Path(directory)
            environment = {
                "RTF_NATIVE_QUAL_OUTPUT": str(root / "output"),
                "RTF_BEFORE_SOURCE": str(root / "before"),
                "RTF_AFTER_SOURCE": str(root / "after"),
                "RTF_TARGET_ROOT": str(root / "targets"),
                "RTF_NATIVE_TMP": str(root / "tmp"),
                "RTF_BASELINE_REVISION": "0" * 40,
                "RTF_CANDIDATE_REVISION": "1" * 40,
                "LITCHI_COORD_BIN": str(root / "coord-bin"),
                "LITCHI_NATIVE_QUAL_ROOT": str(ROOT),
            }
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                manifest = ROOT / "tools/native_qualification/manifests/legacy-rtf.json"
                for receipt in (
                    root / "before" / "receipt.json",
                    root / "targets" / "before" / "receipt.json",
                    root / "output" / "receipt.json",
                ):
                    os.environ["RTF_NATIVE_PREPARATION_RECEIPT"] = str(receipt)
                    with self.assertRaisesRegex(runner.RunnerError, "outside|pairwise disjoint"):
                        runner.load_manifest(manifest)
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

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
                lock = source / "tools/perf-baseline/Cargo.lock"
                probe = source / "tools/perf-baseline/examples/refusal_contract_bench.rs"
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
            manifest_path = root / "legacy-rtf-final.json"
            manifest_path.write_bytes(
                (ROOT / "tools/native_qualification/manifests/legacy-rtf.json").read_bytes()
            )
            with self.assertRaisesRegex(validator.ValidationError, "manifest path is not source-bound"):
                validator.validate_directory(
                    output,
                    "legacy-rtf",
                    source_heads["before"],
                    source_heads["after"],
                    targets["before"],
                    targets["after"],
                    manifest_path,
                )

    def test_validator_binds_external_manifest_hash_and_lane_before_receipt(self) -> None:
        with tempfile.TemporaryDirectory(prefix="legacy-native-manifest-binding-") as directory:
            root = Path(directory)
            manifest_path = root / "legacy-rtf-final.json"
            manifest_path.write_bytes(
                (ROOT / "tools/native_qualification/manifests/legacy-rtf.json").read_bytes()
            )
            output = root / "output"
            output.mkdir()
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
                "manifest": {
                    "path": str(manifest_path),
                    "sha256": hashlib.sha256(manifest_path.read_bytes()).hexdigest(),
                },
                "identities": {},
            }
            receipt_path = output / "build-receipt.json"

            receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "runner path is not source-bound"):
                validator._validate_build_receipt(
                    output,
                    "legacy-rtf",
                    validator.CONTRACTS["legacy-rtf"],
                    validator.CONTRACTS["legacy-rtf"]["baseline"],
                    "1" * 40,
                    root / "before-target",
                    root / "after-target",
                    manifest_path,
                )

            receipt["manifest"]["path"] = str(ROOT / "tools/native_qualification/manifests/legacy-rtf.json")
            receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "manifest path is not source-bound"):
                validator._validate_build_receipt(
                    output,
                    "legacy-rtf",
                    validator.CONTRACTS["legacy-rtf"],
                    validator.CONTRACTS["legacy-rtf"]["baseline"],
                    "1" * 40,
                    root / "before-target",
                    root / "after-target",
                    manifest_path,
                )

            receipt["manifest"]["path"] = str(manifest_path)
            manifest_path.write_text(
                manifest_path.read_text(encoding="utf-8") + "\n",
                encoding="utf-8",
            )
            receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "manifest hash is not source-bound"):
                validator._validate_build_receipt(
                    output,
                    "legacy-rtf",
                    validator.CONTRACTS["legacy-rtf"],
                    validator.CONTRACTS["legacy-rtf"]["baseline"],
                    "1" * 40,
                    root / "before-target",
                    root / "after-target",
                    manifest_path,
                )

            manifest_path.write_bytes(
                (ROOT / "tools/native_qualification/manifests/legacy-rtf.json").read_bytes()
            )
            manifest_path.write_text(
                manifest_path.read_text(encoding="utf-8").replace('"lane": "legacy-rtf"', '"lane": "legacy-xls-properties"'),
                encoding="utf-8",
            )
            receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaisesRegex(validator.ValidationError, "qualification manifest lane mismatch"):
                validator._validate_build_receipt(
                    output,
                    "legacy-rtf",
                    validator.CONTRACTS["legacy-rtf"],
                    validator.CONTRACTS["legacy-rtf"]["baseline"],
                    "1" * 40,
                    root / "before-target",
                    root / "after-target",
                    manifest_path,
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
