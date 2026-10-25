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
            identities = {"before": {}, "after": {}}
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
                rows.append({
                    **route.json(),
                    "status": route.status,
                    "evidence": {"source_bindings": bindings, "output_artifact": artifact},
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
                "output_root": str(output_root.resolve()),
                "source_artifacts": identities,
                "tooling": tooling,
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
                probes=(artifact(policy.PPT_PROBE_PATH), artifact(policy.LEGACY_PROBE_PATH)),
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
