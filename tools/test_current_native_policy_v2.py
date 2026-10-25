"""Synthetic tests for the current-source native policy v2 boundary."""

from __future__ import annotations

import copy
import hashlib
import io
import inspect
import json
import os
from pathlib import Path
import shutil
import subprocess
import tarfile
import sys
import tempfile
import unittest
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import current_native_policy_v2 as policy  # noqa: E402
from native_qualification import materialize_current_native_adapters as adapters  # noqa: E402
from native_qualification import prepare_current_native_inputs_v2 as preparer  # noqa: E402
from native_qualification import validate_current_native_v2 as validator  # noqa: E402


class CurrentNativePolicyV2Tests(unittest.TestCase):
    def test_candidate_is_pending_and_execution_fails_closed(self) -> None:
        self.assertEqual(policy.TRUSTED_POLICY.status, "incomplete")
        self.assertEqual(len(policy.SOURCE_TRANSITIONS), 3)
        with self.assertRaises(policy.PolicyIncomplete):
            policy.TRUSTED_POLICY.require_execution_ready()
        with self.assertRaises(policy.PolicyIncomplete):
            policy.make_ppt_steps_v2(Path("/tmp/source"), Path("/tmp/state"), Path("/tmp/target"))

    def test_hints_are_not_accepted_as_hashes(self) -> None:
        for hint in policy.CANDIDATE_HINTS.values():
            self.assertIsNone(policy.SHA256.fullmatch(hint))
        self.assertNotIn("expected_sha256", inspect.signature(policy.validate_lock).parameters)
        self.assertNotIn("candidate_lock_sha256", inspect.signature(policy.make_ppt_steps_v2).parameters)

    def test_manifests_match_the_fixed_route_contract(self) -> None:
        for name, lane in (("ppt-current-v2.json", "ppt-auxiliary"), ("legacy-current-v2.json", "legacy-current")):
            path = ROOT / "tools/native_qualification/manifests" / name
            value = json.loads(path.read_text(encoding="utf-8"))
            policy.validate_routes(lane, value["routes"])
            self.assertEqual(value["policy_schema"], policy.POLICY_SCHEMA)

        value = json.loads((ROOT / "tools/native_qualification/manifests/ppt-current-v2.json").read_text())
        tampered = copy.deepcopy(value)
        tampered["routes"][0]["output"] = "target/release/other"
        with self.assertRaises(policy.PolicyMismatch):
            policy.validate_routes("ppt-auxiliary", tampered["routes"])
        refused = copy.deepcopy(value)
        for route in refused["routes"]:
            route["status"] = "refused"
        with self.assertRaises(policy.PolicyMismatch):
            policy.validate_routes("ppt-auxiliary", refused["routes"])

    def test_pending_preparation_is_diagnostic_only(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before = self._fake_source(parent / "before", "before")
            after = self._fake_source(parent / "after", "after")
            output = parent / "receipt.json"
            receipt = preparer.prepare(
                lane="ppt-auxiliary",
                manifest=manifest,
                before=before,
                after=after,
                output=output,
            )
            self.assertEqual(receipt["status"], "blocked")
            self.assertFalse(receipt["execution_admitted"])
            self.assertEqual(receipt["policy_sha256"], policy.TRUSTED_POLICY.digest())
            self.assertEqual(json.loads(output.read_text()), receipt)
            with self.assertRaises(policy.PolicyIncomplete):
                validator.validate_preparation(output)

    def test_preparation_rejects_caller_route_overlay(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        value = json.loads(manifest.read_text(encoding="utf-8"))
        value["routes"][1]["program"] = "caller-selected"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "manifest.json"
            path.write_bytes((json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode())
            before = self._fake_source(Path(directory) / "before", "before")
            after = self._fake_source(Path(directory) / "after", "after")
            with self.assertRaises(preparer.PreparationError):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=path,
                    before=before,
                    after=after,
                    output=Path(directory) / "receipt.json",
                )

    def test_sparse_omission_is_rejected_even_with_the_same_head(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "source", "sparse")
            (source / "tracked.txt").write_text("tracked payload\n")
            subprocess.run(["git", "add", "tracked.txt"], cwd=source, check=True)
            subprocess.run(["git", "commit", "-q", "-m", "tracked"], cwd=source, check=True)
            subprocess.run(["git", "update-index", "--skip-worktree", "tracked.txt"], cwd=source, check=True)
            (source / "tracked.txt").unlink()
            with self.assertRaises(policy.PolicyMismatch):
                policy.observe_source(source, "candidate")

    def test_ignored_overlay_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "source", "ignored")
            (source / ".gitignore").write_text("overlay.bin\n")
            subprocess.run(["git", "add", ".gitignore"], cwd=source, check=True)
            subprocess.run(["git", "commit", "-q", "-m", "ignore"], cwd=source, check=True)
            (source / "overlay.bin").write_bytes(b"unreviewed overlay")
            with self.assertRaises(policy.PolicyMismatch):
                policy.observe_source(source, "candidate")

    def test_exact_pinned_root_lock_is_the_only_allowed_ignored_overlay(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "source", "ignored-lock")
            (source / ".gitignore").write_text("Cargo.lock\n")
            subprocess.run(["git", "add", ".gitignore"], cwd=source, check=True)
            subprocess.run(["git", "commit", "-q", "-m", "ignore-lock"], cwd=source, check=True)
            lock = source / "Cargo.lock"
            lock.write_bytes(b"pinned lock\n")
            digest = hashlib.sha256(lock.read_bytes()).hexdigest()
            entries = policy._git_tree_entries(source)
            identity = policy._tracked_materialization(
                source,
                entries,
                (policy.ArtifactPin("Cargo.lock", digest, lock.stat().st_size),),
            )
            self.assertGreaterEqual(identity["files"], 2)

    def test_tracked_symlink_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "source", "symlink")
            (source / "link").symlink_to("Cargo.toml")
            subprocess.run(["git", "add", "link"], cwd=source, check=True)
            subprocess.run(["git", "commit", "-q", "-m", "symlink"], cwd=source, check=True)
            with self.assertRaises(policy.PolicyMismatch):
                policy.observe_source(source, "candidate")

    def test_preparation_rejects_ancestor_and_descendant_roots(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before = self._fake_source(parent / "before", "ancestor")
            nested = before / "nested"
            nested.mkdir()
            with self.assertRaises(preparer.PreparationError):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=before,
                    after=nested,
                    output=parent / "receipt.json",
                )

    def test_ppt_plan_rejects_source_output_archive_overlap(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "source", "plan")
            with self.assertRaises(policy.PolicyMismatch):
                policy._check_plan_paths(source, source / "state", Path(directory) / "target", None)
            with self.assertRaises(policy.PolicyMismatch):
                policy._check_plan_paths(source, Path(directory) / "state", Path(directory) / "target", source / "archive.tar")
            with self.assertRaises(policy.PolicyMismatch):
                policy._check_plan_paths(ROOT, Path(directory) / "state", Path(directory) / "target", None)

    def test_preparation_rejects_protocol_checkout_as_source(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            after = self._fake_source(parent / "after", "protocol-boundary")
            with self.assertRaises(preparer.PreparationError):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=ROOT,
                    after=after,
                    output=parent / "receipt.json",
                )

    def test_output_artifact_must_be_the_policy_path_under_the_prepared_root(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "outputs"
            expected = policy.PPT_ROUTES[0]
            output = root / expected.output
            output.parent.mkdir(parents=True)
            output.write_bytes(b"bound output")
            identity = {
                "path": str(output.resolve()),
                "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
                "bytes": output.stat().st_size,
                "relative_path": expected.output,
            }
            validator._validate_output_artifact({"output_artifact": identity}, expected, root.resolve(), "output")
            unrelated = Path(directory) / "unrelated"
            unrelated.write_bytes(b"bound output")
            forged = dict(identity)
            forged["path"] = str(unrelated)
            with self.assertRaises(validator.ValidationError):
                validator._validate_output_artifact({"output_artifact": forged}, expected, root.resolve(), "output")
            escaped_root = Path(directory) / "escaped-root"
            escaped_root.mkdir()
            outside = Path(directory) / "outside"
            outside_output = outside / expected.output
            outside_output.parent.mkdir(parents=True)
            outside_output.write_bytes(b"escaped output")
            (escaped_root / "target").symlink_to(outside / "target", target_is_directory=True)
            escaped_artifact = {
                "path": str(outside_output),
                "sha256": hashlib.sha256(outside_output.read_bytes()).hexdigest(),
                "bytes": outside_output.stat().st_size,
                "relative_path": expected.output,
            }
            with self.assertRaises(validator.ValidationError):
                validator._validate_output_artifact(
                    {"output_artifact": escaped_artifact}, expected, escaped_root.resolve(), "output"
                )

    def test_preparation_requires_a_fresh_output_root(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before = self._fake_source(parent / "before", "fresh-before")
            after = self._fake_source(parent / "after", "fresh-after")
            output_root = parent / "outputs"
            receipt = preparer.prepare(
                lane="ppt-auxiliary",
                manifest=manifest,
                before=before,
                after=after,
                output=parent / "receipt.json",
                output_root=output_root,
            )
            self.assertEqual(receipt["output_root_initial"]["entries"], [])
            self.assertEqual(list(output_root.iterdir()), [])
            with self.assertRaises(preparer.PreparationError):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=before,
                    after=after,
                    output=parent / "second-receipt.json",
                    output_root=output_root,
                )

    def test_execution_receipt_path_tampering_is_rejected(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        for forged_path in ("/tmp/escape.json", "../escape.json"):
            with tempfile.TemporaryDirectory() as directory:
                parent = Path(directory)
                before = self._fake_source(parent / "before", "receipt-before")
                after = self._fake_source(parent / "after", "receipt-after")
                receipt_path = parent / "preparation.json"
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=before,
                    after=after,
                    output=receipt_path,
                    output_root=parent / "outputs",
                )
                receipt = json.loads(receipt_path.read_text())
                receipt["execution_receipt_relative"] = forged_path
                receipt_path.write_bytes((json.dumps(receipt, sort_keys=True, separators=(",", ":")) + "\n").encode())
                with self.assertRaises(validator.ValidationError):
                    validator.validate_preparation(receipt_path)

    def test_final_qualification_rejects_all_refused_rows(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            output_root = parent / "outputs"
            output_root.mkdir()
            identities = {
                "before": {"source_root": str(parent / "before")},
                "after": {"source_root": str(parent / "after")},
            }
            binary_provenance = {
                "before": [
                    {"path": str(parent / "before-bin"), "sha256": "1" * 64, "bytes": 1},
                    {"path": str(parent / "before-correctness-bin"), "sha256": "3" * 64, "bytes": 1},
                ],
                "after": [
                    {"path": str(parent / "after-bin"), "sha256": "2" * 64, "bytes": 1},
                    {"path": str(parent / "after-correctness-bin"), "sha256": "4" * 64, "bytes": 1},
                ],
            }
            tooling = {"test": "tooling"}
            prepared = {
                "receipt": {
                    "identities": identities,
                    "tooling": tooling,
                    "execution_receipt_relative": "execution-receipt.json",
                },
                "receipt_sha256": "preparation",
                "lane": "ppt-auxiliary",
                "output_root": output_root.resolve(),
            }
            routes = [route.json() for route in policy.PPT_ROUTES]
            bindings = {
                side: hashlib.sha256(validator._canonical(identities[side])).hexdigest()
                for side in ("before", "after")
            }
            for route in policy.PPT_ROUTES:
                output = output_root / route.output
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_bytes(route.route.encode())
            rows = []
            for route in policy.PPT_ROUTES:
                output = output_root / route.output
                artifact = {
                    "path": str(output.resolve()),
                    "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
                    "bytes": output.stat().st_size,
                    "relative_path": route.output,
                }
                abba = []
                for index, leg in enumerate(policy.ABBA_ORDER):
                    relative_leg = f".abba/{route.route.replace('/', '_')}/b1-{index}-{leg}.stream"
                    leg_path = output_root / relative_leg
                    leg_path.parent.mkdir(parents=True, exist_ok=True)
                    leg_path.write_bytes(f"{route.route}/{leg}".encode())
                    abba.append({
                        "route": route.route,
                        "role": leg,
                        "index": index,
                        "block": 1,
                        "returncode": 0,
                        "source_binding": bindings[leg] if leg in bindings else bindings["before"],
                        "cwd": identities[leg]["source_root"],
                        "argv": ["cargo", "build", "--release"] if route.kind == "build" else [binary_provenance[leg][1 if route.program == "scoped_patch_correctness" else 0]["path"]],
                        "binary": binary_provenance[leg][1 if route.program == "scoped_patch_correctness" else 0],
                        "output_artifact": {
                            "path": str(leg_path.resolve()),
                            "sha256": hashlib.sha256(leg_path.read_bytes()).hexdigest(),
                            "bytes": leg_path.stat().st_size,
                            "relative_path": relative_leg,
                        },
                    })
                rows.append({
                    **route.json(),
                    "status": route.status,
                    "evidence": {
                        "source_bindings": bindings,
                        "binary_provenance": binary_provenance,
                        "output_artifact": artifact,
                        "abba": abba,
                    },
                })
            execution = {
                "schema": "litchi.current-native-execution.v2",
                "status": "completed",
                "policy_sha256": policy.TRUSTED_POLICY.digest(),
                "preparation_receipt_sha256": "preparation",
                "lane": "ppt-auxiliary",
                "output_root": str(output_root.resolve()),
                "routes": routes,
                "source_artifacts": identities,
                "tooling": tooling,
                "abba_order": list(policy.ABBA_ORDER),
                "abba_blocks": 1,
                "adapter_features": {"before": ["all-features"], "after": ["all-features"]},
                "binary_provenance": binary_provenance,
                "runner": validator._file_identity(ROOT / "tools/native_qualification_runner.py", "runner"),
                "route_receipts": rows,
            }
            execution_path = output_root / "execution-receipt.json"
            execution_path.write_bytes((json.dumps(execution, sort_keys=True, separators=(",", ":")) + "\n").encode())
            qualification = {
                "schema": policy.QUALIFICATION_SCHEMA,
                "policy_schema": policy.POLICY_SCHEMA,
                "policy_sha256": policy.TRUSTED_POLICY.digest(),
                "preparation_receipt_sha256": "preparation",
                "lane": "ppt-auxiliary",
                "abba_order": list(policy.ABBA_ORDER),
                "abba_blocks": 1,
                "output_root": str(output_root.resolve()),
                "source_artifacts": identities,
                "tooling": tooling,
                "binary_provenance": binary_provenance,
                "routes": routes,
                "results": rows,
                "execution_receipt": {
                    "path": str(execution_path.resolve()),
                    "sha256": hashlib.sha256(execution_path.read_bytes()).hexdigest(),
                    "bytes": execution_path.stat().st_size,
                    "relative_path": "execution-receipt.json",
                },
            }
            qualification_path = parent / "qualification.json"
            qualification_path.write_bytes((json.dumps(qualification, sort_keys=True, separators=(",", ":")) + "\n").encode())
            with mock.patch.object(validator, "validate_preparation", return_value=prepared):
                validator.validate_qualification(qualification_path, parent / "preparation.json")
                forged = copy.deepcopy(qualification)
                for row in forged["results"]:
                    row["status"] = "refused"
                qualification_path.write_bytes((json.dumps(forged, sort_keys=True, separators=(",", ":")) + "\n").encode())
                with self.assertRaises(validator.ValidationError):
                    validator.validate_qualification(qualification_path, parent / "preparation.json")

    def test_manifest_mutation_is_checked_before_pending_admission(self) -> None:
        source_manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            manifest = parent / "manifest.json"
            manifest.write_bytes(source_manifest.read_bytes())
            before = self._fake_source(parent / "before", "manifest-before")
            after = self._fake_source(parent / "after", "manifest-after")
            receipt_path = parent / "receipt.json"
            preparer.prepare(
                lane="ppt-auxiliary",
                manifest=manifest,
                before=before,
                after=after,
                output=receipt_path,
            )
            value = json.loads(manifest.read_text())
            value["status"] = "caller-rewritten"
            manifest.write_bytes((json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode())
            with self.assertRaises(validator.ValidationError):
                validator.validate_preparation(receipt_path)

    def test_tooling_identity_mutation_is_checked_before_pending_admission(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before = self._fake_source(parent / "before", "tooling-before")
            after = self._fake_source(parent / "after", "tooling-after")
            receipt_path = parent / "receipt.json"
            preparer.prepare(
                lane="ppt-auxiliary",
                manifest=manifest,
                before=before,
                after=after,
                output=receipt_path,
            )
            value = json.loads(receipt_path.read_text())
            value["tooling"]["validator"]["sha256"] = "0" * 64
            receipt_path.write_bytes((json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode())
            with self.assertRaises(validator.ValidationError):
                validator.validate_preparation(receipt_path)

    def test_fabricated_qualification_cannot_bypass_pending_policy(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            qualification = Path(directory) / "qualification.json"
            qualification.write_bytes(b"{}\n")
            with self.assertRaises(validator.ValidationError):
                validator.validate_qualification(qualification, qualification)

    def test_ready_pipeline_binds_three_roles_and_both_consumers(self) -> None:
        """Exercise the admitted preparation path with real source/archive proofs."""

        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before, before_archive = self._complete_source(parent / "before", "before", "before.tar")
            after, after_archive = self._complete_source(parent / "after", "after", "after.tar")
            trusted = self._complete_policy(before, before_archive, after, after_archive)
            manifest = self._manifest_for_policy(parent / "manifest.json", "ppt-auxiliary", trusted)
            output_root = parent / "outputs"
            receipt_path = parent / "preparation.json"
            with mock.patch.object(policy, "TRUSTED_POLICY", trusted), \
                mock.patch.object(preparer, "TRUSTED_POLICY", trusted), \
                mock.patch.object(validator, "TRUSTED_POLICY", trusted):
                receipt = preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=before,
                    after=after,
                    before_archive=before_archive,
                    after_archive=after_archive,
                    output_root=output_root,
                    output=receipt_path,
                )
                self.assertTrue(receipt["execution_admitted"])
                self.assertEqual(set(receipt["identities"]["after"]["artifacts"]["role_locks"]), set(policy.ROLE_NAMES))
                admitted = validator.validate_preparation(receipt_path)
                self.assertEqual(admitted["lane"], "ppt-auxiliary")
                ppt_steps = policy.make_ppt_steps_v2(after, parent / "ppt-state", parent / "ppt-target", source_archive=after_archive)
                legacy_steps = policy.make_legacy_steps_v2(after, parent / "legacy-state", parent / "legacy-target", source_archive=after_archive)
                self.assertEqual(ppt_steps[0]["argv"][2:4], ["guard-role-lock", "workspace"])
                self.assertEqual(legacy_steps[0]["argv"][2:4], ["guard-role-lock", "perf"])

    def test_injected_executor_writes_and_consumes_v2_abba_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before, before_archive = self._complete_source(parent / "before", "before-exec", "before.tar")
            after, after_archive = self._complete_source(parent / "after", "after-exec", "after.tar")
            trusted = self._complete_policy(before, before_archive, after, after_archive)
            manifest = self._manifest_for_policy(parent / "manifest.json", "ppt-auxiliary", trusted)
            preparation_path = parent / "preparation.json"
            output_root = parent / "outputs"
            qualification_path = parent / "qualification.json"
            calls: list[dict[str, object]] = []

            def fake_executor(step: dict[str, object]) -> dict[str, object]:
                calls.append(step.copy())
                if step.get("produces_binaries"):
                    identities = []
                    for raw_path in step["binary_paths"]:
                        path = Path(str(raw_path))
                        path.parent.mkdir(parents=True, exist_ok=True)
                        path.write_bytes(f"{step['side']}:{path.name}".encode())
                        identities.append({
                            "path": str(path.resolve()),
                            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                            "bytes": path.stat().st_size,
                        })
                    return {"returncode": 0, "binary_identities": identities}
                route = step.get("route", step.get("name", "build"))
                role = step.get("abba_role", "build")
                if route == "ppt-scoped-correctness":
                    fixture = Path(str(step["fixture"]))
                    source = fixture.read_bytes()
                    payload = {
                        "mode": "scoped_patch_correctness",
                        "fixture": str(fixture),
                        "media_id": 1,
                        "requested_external_media_path": "clip-0000.avi",
                        "before_external_media_path": None,
                        "external_media_path": "clip-0000.avi",
                        "source_len": len(source),
                        "target_len": len(source) + 1,
                        "source_sha256": hashlib.sha256(source).hexdigest(),
                        "target_sha256": "a" * 64,
                        "replay_target_sha256": "a" * 64,
                        "replay_matches_target": True,
                        "inverse_source_sha256": hashlib.sha256(source).hexdigest(),
                        "inverse_matches_source": True,
                    }
                    return {"returncode": 0, "output_bytes": json.dumps(payload, sort_keys=True).encode() + b"\n"}
                if route == "ppt-scoped-correctness-validate":
                    payload = {
                        "route": route,
                        "status": "passed",
                        "target_sha256": "a" * 64,
                        "replay_target_sha256": "a" * 64,
                        "replay_matches_target": True,
                        "external_media_path": "clip-0000.avi",
                        "inverse_matches_source": True,
                    }
                    return {"returncode": 0, "output_bytes": json.dumps(payload, sort_keys=True).encode() + b"\n"}
                if route == "ppt-scoped-correctness-suite":
                    payload = {"route": route, "status": "ok"}
                    return {"returncode": 0, "output_bytes": json.dumps(payload, sort_keys=True).encode() + b"\n"}
                return {
                    "returncode": 0,
                    "output_bytes": json.dumps({"route": route, "role": role}, sort_keys=True).encode() + b"\n",
                }

            with mock.patch.object(policy, "TRUSTED_POLICY", trusted), \
                mock.patch.object(preparer, "TRUSTED_POLICY", trusted), \
                mock.patch.object(validator, "TRUSTED_POLICY", trusted):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=manifest,
                    before=before,
                    after=after,
                    before_archive=before_archive,
                    after_archive=after_archive,
                    output_root=output_root,
                    output=preparation_path,
                )
                result = policy.execute_v2(preparation_path, qualification_path, executor=fake_executor)
                self.assertEqual(result["schema"], policy.QUALIFICATION_SCHEMA)
                self.assertEqual(result["lane"], "ppt-auxiliary")
                self.assertEqual(len(calls), 12)
                self.assertEqual(calls[0]["name"], "ppt-scoped-build")
                self.assertEqual(calls[0]["side"], "before")
                self.assertEqual(calls[2]["name"], "ppt-scoped-build")
                self.assertEqual(calls[2]["side"], "after")
                self.assertNotEqual(calls[0]["cwd"], calls[2]["cwd"])
                self.assertNotEqual(calls[0]["binary_paths"], calls[2]["binary_paths"])
                route_calls = [step for step in calls if step.get("abba_role")]
                self.assertTrue(all(step["cwd"] == (str(before) if step["abba_role"] == "before" else str(after)) for step in route_calls))
                self.assertTrue(all(step["binary"] in step["argv"] for step in route_calls if step["kind"] == "argv" and step["route"] == "ppt-scoped-correctness"))
                value = json.loads(qualification_path.read_text(encoding="utf-8"))
                self.assertEqual([row["route"] for row in value["results"]], [route.route for route in policy.PPT_ROUTES])
                self.assertTrue(all(len(row["evidence"]["abba"]) == 4 for row in value["results"]))
                value["results"][0]["evidence"]["abba"][0]["role"] = "after"
                qualification_path.write_bytes(policy._canonical(value))
                with self.assertRaises(validator.ValidationError):
                    validator.validate_qualification(qualification_path, preparation_path)

    def test_injected_executor_preserves_eight_legacy_route_streams(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before, before_archive = self._complete_source(parent / "before", "before-legacy", "before.tar")
            after, after_archive = self._complete_source(parent / "after", "after-legacy", "after.tar")
            trusted = self._complete_policy(before, before_archive, after, after_archive)
            manifest = self._manifest_for_policy(parent / "manifest.json", "legacy-current", trusted)
            preparation_path = parent / "preparation.json"
            output_root = parent / "outputs"
            qualification_path = parent / "qualification.json"
            calls: list[dict[str, object]] = []

            def fake_executor(step: dict[str, object]) -> dict[str, object]:
                calls.append(step)
                if step.get("produces_binaries"):
                    identities = []
                    for raw_path in step["binary_paths"]:
                        path = Path(str(raw_path))
                        path.parent.mkdir(parents=True, exist_ok=True)
                        path.write_bytes(f"{step['side']}:{path.name}".encode())
                        identities.append({
                            "path": str(path.resolve()),
                            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                            "bytes": path.stat().st_size,
                        })
                    return {"returncode": 0, "binary_identities": identities}
                rows = []
                for fixed_route in policy.LEGACY_CURRENT_ROUTES:
                    rows.append({
                        "scenario": fixed_route.route.rsplit("/", 1)[-1],
                        "median_us": 1.0,
                        "workload": "rtf-refusal-contract-v1",
                        "outcome": "ok" if fixed_route.status == "ok" else "refused",
                        "fidelity": "exact" if fixed_route.status == "ok" else "unchanged",
                        **(
                            {
                                "refusal_kind": "protected" if fixed_route.route.endswith("protected_commit") else "unsupported_content",
                                "operation": "edit",
                                "reason": "document protection refuses changed publication" if fixed_route.route.endswith("protected_commit") else "underline edits require non-empty text within one paragraph",
                            }
                            if fixed_route.status == "refused" else {}
                        ),
                    })
                return {
                    "returncode": 0,
                    "output_bytes": b"".join(
                        (json.dumps(row, sort_keys=True).encode() + b"\n") for row in rows
                    ),
                }

            with mock.patch.object(policy, "TRUSTED_POLICY", trusted), \
                mock.patch.object(preparer, "TRUSTED_POLICY", trusted), \
                mock.patch.object(validator, "TRUSTED_POLICY", trusted):
                preparer.prepare(
                    lane="legacy-current",
                    manifest=manifest,
                    before=before,
                    after=after,
                    before_archive=before_archive,
                    after_archive=after_archive,
                    output_root=output_root,
                    output=preparation_path,
                )
                policy.execute_v2(preparation_path, qualification_path, executor=fake_executor)
                value = json.loads(qualification_path.read_text(encoding="utf-8"))
                self.assertEqual(len(value["results"]), 8)
                self.assertEqual(
                    [row["route"] for row in value["results"]],
                    [route.route for route in policy.LEGACY_CURRENT_ROUTES],
                )
                self.assertTrue(all(len(row["evidence"]["abba"]) == 20 for row in value["results"]))
                route_calls = [step for step in calls if step.get("route")]
                self.assertEqual(len(route_calls), 8 * 5 * len(policy.ABBA_ORDER))
                self.assertTrue(all("--check-only" not in step["argv"] for step in route_calls))
                self.assertEqual({step["side"] for step in route_calls}, {"before", "after"})
                self.assertEqual({step["adapter_features"][0] for step in route_calls}, {"rtf-legacy-api", "rtf-current-api"})

    def test_route_payload_rejects_untyped_bytes_and_requires_fixed_legacy_fields(self) -> None:
        route = policy.LEGACY_CURRENT_ROUTES[0]
        with self.assertRaises(policy.PolicyMismatch):
            policy._v2_result_payload({"output_bytes": b"arbitrary native bytes"}, route, Path("/tmp"))
        with self.assertRaises(policy.PolicyMismatch):
            policy._v2_result_payload({"stdout": b"not-json\n"}, route, Path("/tmp"))
        with self.assertRaises(policy.PolicyMismatch):
            policy._v2_result_payload(
                {"output_bytes": json.dumps({"scenario": "open_1k", "median_us": 1}).encode()},
                route,
                Path("/tmp"),
            )

    def test_reviewed_f2_adapter_sources_are_fixed_and_external(self) -> None:
        baseline = policy.TRUSTED_POLICY.baseline
        self.assertEqual(
            {item.relative_path for item in baseline.adapter_overlays},
            {
                "tools/perf-baseline/Cargo.toml",
                policy.LEGACY_PROBE_PATH,
                policy.PPT_CORRECTNESS_PATH,
            },
        )
        for item in baseline.adapter_overlays:
            identity = policy._validate_adapter_source(item, f"test adapter {item.relative_path}")
            self.assertEqual(identity["sha256"], item.sha256)
            self.assertEqual(identity["bytes"], item.bytes)
            self.assertIn("tools/native_qualification/adapters/", identity["source_path"])
        correctness = baseline.adapter_for(policy.PPT_CORRECTNESS_PATH)
        self.assertIsNotNone(correctness)
        self.assertEqual(correctness.sha256, "3bb1deaa67f64e8cb70f7dc2cf05ad41803fa8edd2f73b492991555d260579ab")

    def test_f2_correctness_receipt_requires_forward_replay(self) -> None:
        route = policy.PPT_ROUTES[2]
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            fixture = parent / "video.ppt"
            fixture.write_bytes(b"f2 correctness fixture")
            source_hash = hashlib.sha256(fixture.read_bytes()).hexdigest()
            payload = {
                "mode": "scoped_patch_correctness",
                "fixture": str(fixture),
                "media_id": 1,
                "requested_external_media_path": "clip-0000.avi",
                "before_external_media_path": None,
                "external_media_path": "clip-0000.avi",
                "source_len": fixture.stat().st_size,
                "target_len": fixture.stat().st_size + 1,
                "source_sha256": source_hash,
                "target_sha256": "a" * 64,
                "replay_target_sha256": "a" * 64,
                "replay_matches_target": True,
                "inverse_source_sha256": source_hash,
                "inverse_matches_source": True,
            }
            bound = policy._v2_result_payload(
                {"output_bytes": json.dumps(payload).encode()},
                route,
                parent,
                expected_fixture=fixture,
            )
            self.assertEqual(json.loads(bound), payload)
            tagged = dict(payload)
            tagged["route"] = route.route
            tagged_bound = policy._v2_result_payload(
                {"output_bytes": json.dumps(tagged).encode()},
                route,
                parent,
                expected_fixture=fixture,
            )
            self.assertEqual(json.loads(tagged_bound), tagged)
            for field, value in (("media_id", 2), ("external_media_path", "wrong")):
                tampered = dict(payload)
                tampered[field] = value
                with self.assertRaises(policy.PolicyMismatch):
                    policy._v2_result_payload(
                        {"output_bytes": json.dumps(tampered).encode()},
                        route,
                        parent,
                        expected_fixture=fixture,
                    )
            missing_replay = dict(payload)
            del missing_replay["replay_target_sha256"]
            with self.assertRaises(policy.PolicyMismatch):
                policy._v2_result_payload(
                    {"output_bytes": json.dumps(missing_replay).encode()},
                    route,
                    parent,
                    expected_fixture=fixture,
                )

            for tagged_route, tagged in (
                (route, {"route": "ppt-scoped-correctness"}),
                (policy.PPT_ROUTES[1], {"route": "ppt-scoped-correctness-suite"}),
            ):
                with self.assertRaises(policy.PolicyMismatch):
                    policy._v2_result_payload(
                        {"output_bytes": json.dumps(tagged).encode()},
                        tagged_route,
                        parent,
                        expected_fixture=fixture,
                    )

    def test_f2_correctness_validator_requires_reviewed_forward_replay(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            fixture = parent / "video.ppt"
            fixture.write_bytes(b"reviewed f2 fixture")
            source_hash = hashlib.sha256(fixture.read_bytes()).hexdigest()
            receipt = parent / "correctness.json"
            receipt.write_bytes(
                json.dumps(
                    {
                        "mode": "scoped_patch_correctness",
                        "fixture": str(fixture),
                        "media_id": 1,
                        "requested_external_media_path": "clip-0000.avi",
                        "before_external_media_path": None,
                        "external_media_path": "clip-0000.avi",
                        "source_len": fixture.stat().st_size,
                        "target_len": fixture.stat().st_size + 1,
                        "source_sha256": source_hash,
                        "target_sha256": "b" * 64,
                        "replay_target_sha256": "b" * 64,
                        "replay_matches_target": True,
                        "inverse_source_sha256": source_hash,
                        "inverse_matches_source": True,
                    }
                ).encode()
            )
            completed = subprocess.run(
                [sys.executable, str(ROOT / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"), str(receipt), str(fixture)],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr)
            self.assertIn('"status": "passed"', completed.stdout)

            row = json.loads(receipt.read_text(encoding="utf-8"))
            del row["replay_target_sha256"]
            del row["replay_matches_target"]
            receipt.write_bytes(json.dumps(row).encode())
            failed = subprocess.run(
                [sys.executable, str(ROOT / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"), str(receipt), str(fixture)],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            self.assertNotEqual(failed.returncode, 0)

    def test_adapter_materializer_rejects_non_f2_checkout(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = self._fake_source(Path(directory) / "not-f2", "not-f2")
            with self.assertRaises(adapters.MaterializationError):
                adapters.materialize_baseline_adapters(source)

    def test_execution_rejects_disallowed_qualification_overlap_before_child(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before, before_archive = self._complete_source(parent / "before", "overlap-before", "before.tar")
            after, after_archive = self._complete_source(parent / "after", "overlap-after", "after.tar")
            trusted = self._complete_policy(before, before_archive, after, after_archive)
            manifest = self._manifest_for_policy(parent / "manifest.json", "ppt-auxiliary", trusted)
            preparation_path = parent / "preparation.json"
            output_root = parent / "outputs"
            calls: list[dict[str, object]] = []

            def fake_executor(step: dict[str, object]) -> dict[str, object]:
                calls.append(step)
                return {"returncode": 0}

            with mock.patch.object(policy, "TRUSTED_POLICY", trusted), \
                mock.patch.object(preparer, "TRUSTED_POLICY", trusted), \
                mock.patch.object(validator, "TRUSTED_POLICY", trusted):
                preparer.prepare(
                    lane="ppt-auxiliary", manifest=manifest, before=before, after=after,
                    before_archive=before_archive, after_archive=after_archive,
                    output_root=output_root, output=preparation_path,
                )
                with self.assertRaises(policy.PolicyMismatch):
                    policy.execute_v2(preparation_path, before / "qualification.json", executor=fake_executor)
            self.assertEqual(calls, [])

    def test_ready_pipeline_rejects_wrong_source_lock_archive_and_route(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before, before_archive = self._complete_source(parent / "before", "before", "before.tar")
            after, after_archive = self._complete_source(parent / "after", "after", "after.tar")
            wrong, wrong_archive = self._complete_source(parent / "wrong", "wrong", "wrong.tar")
            trusted = self._complete_policy(before, before_archive, after, after_archive)
            manifest = self._manifest_for_policy(parent / "manifest.json", "ppt-auxiliary", trusted)
            with mock.patch.object(policy, "TRUSTED_POLICY", trusted), \
                mock.patch.object(preparer, "TRUSTED_POLICY", trusted), \
                mock.patch.object(validator, "TRUSTED_POLICY", trusted):
                with self.assertRaises(policy.PolicyMismatch):
                    policy.validate_source(wrong, "candidate", wrong_archive)
                with self.assertRaises(policy.PolicyMismatch):
                    policy.validate_lock(parent / "wrong-lock", "candidate", "perf")
                copied_lock = parent / "copied-lock"
                shutil.copyfile(after / policy.ROLE_LOCK_PATHS["perf"], copied_lock)
                with self.assertRaises(policy.PolicyMismatch):
                    policy.validate_lock(copied_lock, "candidate", "perf")
                bad_archive = parent / "bad.tar"
                shutil.copyfile(after_archive, bad_archive)
                with bad_archive.open("ab") as stream:
                    stream.write(b"tampered")
                with self.assertRaises(policy.PolicyMismatch):
                    policy.validate_source(after, "candidate", bad_archive)
                tampered = json.loads(manifest.read_text(encoding="utf-8"))
                tampered["routes"][0]["output"] = "state/caller-selected.json"
                tampered_path = parent / "tampered-manifest.json"
                tampered_path.write_bytes(policy._canonical(tampered))
                with self.assertRaises(preparer.PreparationError):
                    preparer.prepare(
                        lane="ppt-auxiliary",
                        manifest=tampered_path,
                        before=before,
                        after=after,
                        before_archive=before_archive,
                        after_archive=after_archive,
                        output=parent / "tampered-receipt.json",
                    )

    @staticmethod
    def _fake_source(path: Path, marker: str) -> Path:
        path.mkdir(parents=True)
        (path / "Cargo.toml").write_text(f"[package]\nname = \"{marker}\"\nversion = \"0.0.0\"\n")
        subprocess.run(["git", "init", "-q"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.email", "policy-test@example.invalid"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.name", "Policy Test"], cwd=path, check=True)
        subprocess.run(["git", "add", "Cargo.toml"], cwd=path, check=True)
        subprocess.run(["git", "commit", "-q", "-m", "fixture"], cwd=path, check=True)
        return path

    @staticmethod
    def _complete_source(path: Path, marker: str, archive_name: str) -> tuple[Path, Path]:
        files = {
            "Cargo.toml": f"[package]\nname = \"{marker}\"\nversion = \"0.0.0\"\n",
            "Cargo.lock": f"# {marker} workspace lock\n",
            "tools/perf-baseline/Cargo.toml": "[package]\nname = \"perf\"\nversion = \"0.0.0\"\n",
            "tools/perf-baseline/Cargo.lock": f"# {marker} perf lock\n",
            "tools/native-resave/Cargo.toml": "[package]\nname = \"native\"\nversion = \"0.0.0\"\n",
            "tools/native-resave/Cargo.lock": f"# {marker} native lock\n",
            "rust-toolchain.toml": "[toolchain]\nchannel = \"stable\"\n",
            ".cargo/config.toml": "[build]\ntarget-dir = \"target\"\n",
            "crates/litchi-ppt/examples/scoped_patch_probe.rs": f"fn main() {{ println!(\"{marker}\"); }}\n",
            "crates/litchi-ppt/examples/scoped_patch_correctness.rs": f"fn main() {{ println!(\"correctness-{marker}\"); }}\n",
            "tools/perf-baseline/examples/refusal_contract_bench.rs": f"fn main() {{ println!(\"legacy-{marker}\"); }}\n",
            "test-data/unified-ops/0294/video.ppt": f"fixture-{marker}\n",
        }
        for relative, content in files.items():
            target = path / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(content, encoding="utf-8")
        subprocess.run(["git", "init", "-q"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.email", "policy-test@example.invalid"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.name", "Policy Test"], cwd=path, check=True)
        subprocess.run(["git", "add", "."], cwd=path, check=True)
        subprocess.run(["git", "commit", "-q", "-m", "complete fixture"], cwd=path, check=True)
        entries = policy._git_tree_entries(path)
        directories: set[str] = {"snapshot"}
        for item in entries:
            relative = Path(str(item["path"]))
            parent = relative.parent
            while str(parent) not in {"", "."}:
                directories.add("snapshot/" + str(parent).replace("\\", "/"))
                parent = parent.parent
        archive = path.parent / archive_name
        with tarfile.open(archive, "w") as stream:
            root_info = tarfile.TarInfo("snapshot")
            root_info.type = tarfile.DIRTYPE
            root_info.mode = 0o755
            stream.addfile(root_info)
            for directory in sorted(directories - {"snapshot"}):
                info = tarfile.TarInfo(directory)
                info.type = tarfile.DIRTYPE
                info.mode = 0o755
                stream.addfile(info)
            for item in entries:
                relative = str(item["path"])
                payload = (path / relative).read_bytes()
                info = tarfile.TarInfo("snapshot/" + relative)
                info.mode = 0o755 if item["mode"] == "100755" else 0o644
                info.size = len(payload)
                stream.addfile(info, io.BytesIO(payload))
        return path, archive

    @staticmethod
    def _complete_policy(before: Path, before_archive: Path, after: Path, after_archive: Path) -> policy.CurrentNativePolicyV2:
        def pin(source: Path, archive: Path) -> policy.SourcePin:
            entries = policy._git_tree_entries(source)
            listing = policy._git_tree_payload(source)

            def artifact(relative: str) -> policy.ArtifactPin:
                payload = (source / relative).read_bytes()
                return policy.ArtifactPin(relative, hashlib.sha256(payload).hexdigest(), len(payload))

            archive_payload = archive.read_bytes()
            return policy.SourcePin(
                commit=policy._git(source, "rev-parse", "HEAD"),
                git_tree=policy._git(source, "rev-parse", "HEAD^{tree}"),
                tree_sha256=hashlib.sha256(listing).hexdigest(),
                role_locks=tuple(artifact(path) for path in policy.ROLE_LOCK_PATHS.values()),
                source_manifests=tuple(artifact(path) for path in policy.SOURCE_MANIFEST_PATHS),
                toolchain=artifact("rust-toolchain.toml"),
                cargo_config=artifact(".cargo/config.toml"),
                probes=(
                    artifact(policy.PPT_PROBE_PATH),
                    artifact(policy.LEGACY_PROBE_PATH),
                    artifact(policy.PPT_CORRECTNESS_PATH),
                ),
                fixtures=(artifact("test-data/unified-ops/0294/video.ppt"),),
                archive=policy.ArtifactPin(archive.name, hashlib.sha256(archive_payload).hexdigest(), len(archive_payload)),
                target=policy.TARGET,
                profile=policy.PROFILE,
                features=policy.FEATURES,
            )

        return policy.CurrentNativePolicyV2(
            baseline=pin(before, before_archive),
            candidate=pin(after, after_archive),
            ppt_routes=policy.PPT_ROUTES,
            legacy_routes=policy.LEGACY_CURRENT_ROUTES,
        )

    @staticmethod
    def _manifest_for_policy(path: Path, lane: str, trusted: policy.CurrentNativePolicyV2) -> Path:
        name = "ppt-current-v2.json" if lane == "ppt-auxiliary" else "legacy-current-v2.json"
        value = json.loads((ROOT / "tools/native_qualification/manifests" / name).read_text(encoding="utf-8"))
        value["trusted_inputs"] = trusted.manifest_binding()
        path.write_bytes(policy._canonical(value))
        return path


if __name__ == "__main__":
    unittest.main()
