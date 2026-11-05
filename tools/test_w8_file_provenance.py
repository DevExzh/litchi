"""Tests for W8 fixture/dependency provenance and endpoint identity gates."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

from tools import w8_file_provenance as provenance


ROOT = Path(__file__).resolve().parents[1]
PROBE = ROOT / "crates/litchi-ops/examples/file_w8_qualified_probe.rs"
PACKAGE = ROOT / "crates/litchi-ops/Cargo.toml"
TOOLCHAIN = ROOT / "rust-toolchain.toml"
QUALIFICATION_RECIPE = ROOT / "docs/unified-ops/results/0362yh/472b-qualification-commands.sh"
HISTORICAL_RECIPE = ROOT / "docs/unified-ops/results/0362yh/c796-qualification-commands.sh"


class QualificationRecipe(unittest.TestCase):
    def test_current_recipe_binds_release_target_and_build_receipts(self):
        source = QUALIFICATION_RECIPE.read_text(encoding="utf-8")
        self.assertIn("build --offline --locked --target \"$TARGET\" --release", source)
        self.assertEqual(source.count('python3 "$PROVENANCE" build-receipt'), 2)
        self.assertIn('--output "$W8_PREP/build-before.json"', source)
        self.assertIn('--output "$W8_PREP/build-after.json"', source)
        self.assertIn('missing complete before build receipt; run build first', source)
        self.assertIn('missing complete after build receipt; run build first', source)
        self.assertIn('--cargo "$CARGO_BIN"', source)
        self.assertIn('--compiler-artifact "rustc=$RUSTC_BIN"', source)
        self.assertIn('BASE_BUILD_COMMAND="CARGO_BUILD_JOBS=1', source)
        self.assertIn('--target $TARGET --release', source)
        self.assertIn('CANDIDATE_SOURCE=${CANDIDATE_SOURCE:-}', source)
        self.assertIn('candidate source HEAD is not the reviewed 472b commit', source)
        self.assertIn('CANDIDATE_LOCK_SHA256=4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1', source)
        self.assertIn('verify_candidate_archive "$FINAL_OUT/472b-source.tar"', source)
        self.assertIn('ensure_empty_target "$FINAL_TARGET"', source)
        self.assertIn('BASE_VALIDATION_RECEIPT="$W8_PREP/validation-before.json/receipt.json"', source)
        self.assertIn('FINAL_VALIDATION_RECEIPT="$W8_PREP/validation-after.json/receipt.json"', source)
        self.assertIn('--validation-receipt "$BASE_VALIDATION_RECEIPT"', source)
        self.assertIn('--validation-receipt "$FINAL_VALIDATION_RECEIPT"', source)
        self.assertNotIn('"$CARGO" \\\n      build', source)

    def test_c796_recipe_is_not_presented_as_current_executable(self):
        source = HISTORICAL_RECIPE.read_text(encoding="utf-8")
        self.assertIn("HISTORICAL/SUPERSEDED", source)
        self.assertFalse(os.access(HISTORICAL_RECIPE, os.X_OK))

    def test_candidate_source_fence_rejects_wrong_head_dirty_and_wrong_lock(self):
        with tempfile.TemporaryDirectory() as tmp:
            candidate = Path(tmp) / "candidate"
            subprocess.run(["git", "clone", "--shared", "--no-checkout", str(ROOT), str(candidate)], check=True, capture_output=True)
            subprocess.run(["git", "-C", str(candidate), "sparse-checkout", "init", "--no-cone"], check=True, capture_output=True)
            subprocess.run(
                [
                    "git", "-C", str(candidate), "sparse-checkout", "set",
                    "Cargo.lock", "Cargo.toml", "rust-toolchain.toml",
                    "crates/litchi-ops/Cargo.toml", provenance.PROBE_RELATIVE.as_posix(),
                ],
                check=True,
                capture_output=True,
            )
            subprocess.run(
                ["git", "-C", str(candidate), "checkout", "--detach", "472b6705475b073e7fe08f6e11ee17fb1e29f838"],
                check=True,
                capture_output=True,
            )
            (candidate / "Cargo.lock").write_bytes(
                (ROOT / "docs/unified-ops/results/0362yh/images/after/Cargo.lock").read_bytes()
            )

            def check(expected: bool) -> subprocess.CompletedProcess[str]:
                env = os.environ.copy()
                env["CANDIDATE_SOURCE"] = str(candidate)
                completed = subprocess.run(
                    ["bash", "-c", 'source "$1"; verify_candidate_source', "source-check", str(QUALIFICATION_RECIPE)],
                    env=env,
                    text=True,
                    capture_output=True,
                    check=False,
                )
                if expected:
                    self.assertEqual(completed.returncode, 0, completed.stderr)
                else:
                    self.assertNotEqual(completed.returncode, 0)
                return completed

            env = os.environ.copy()
            env["CANDIDATE_SOURCE"] = "relative-candidate"
            relative = subprocess.run(
                ["bash", "-c", 'source "$1"; verify_candidate_source', "source-check", str(QUALIFICATION_RECIPE)],
                env=env,
                text=True,
                capture_output=True,
            )
            self.assertNotEqual(relative.returncode, 0)
            check(True)
            subprocess.run(["git", "-C", str(candidate), "checkout", "--detach", "c796d0acc24c7c6e6fd08980b9e31af6bb86b18b"], check=True, capture_output=True)
            check(False)
            subprocess.run(["git", "-C", str(candidate), "checkout", "--detach", "472b6705475b073e7fe08f6e11ee17fb1e29f838"], check=True, capture_output=True)
            workspace_manifest = candidate / "Cargo.toml"
            original_workspace = workspace_manifest.read_bytes()
            workspace_manifest.write_bytes(original_workspace + b"\n")
            check(False)
            workspace_manifest.write_bytes(original_workspace)
            check(True)
            lock = candidate / "Cargo.lock"
            original_lock = lock.read_bytes()
            lock.write_bytes(original_lock + b"\n")
            check(False)
            lock.write_bytes(original_lock)
            check(True)


class FixtureDescriptor(unittest.TestCase):
    def test_current_probe_has_the_reviewed_workload(self):
        descriptor = provenance.fixture_descriptor(PROBE)
        self.assertEqual(descriptor["route_count"], 109)
        self.assertEqual(tuple(descriptor["families"]), provenance.EXPECTED_FAMILIES)
        self.assertEqual(descriptor["constants"]["paragraphs"], 1_000)
        self.assertEqual(descriptor["constants"]["shapes"], 100)
        self.assertEqual(descriptor["api"]["commit_extraction"], "into_file")
        self.assertEqual(descriptor["probe"]["sha256"], provenance.REVIEWED_PROBE_SHA256)
        self.assertEqual(descriptor["recipe"]["xlsb"]["import_rows"], 100)
        self.assertEqual(descriptor["recipe"]["odp"]["batch"], 100)

    def test_descriptor_bytes_are_deterministic(self):
        descriptor = provenance.fixture_descriptor(PROBE)
        with tempfile.TemporaryDirectory() as tmp:
            first = Path(tmp) / "first.json"
            second = Path(tmp) / "second.json"
            provenance.write_new(first, descriptor)
            provenance.write_new(second, descriptor)
            self.assertEqual(first.read_bytes(), second.read_bytes())
            self.assertEqual(provenance.digest(first), provenance.digest(second))

    def test_changed_workload_constant_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            altered = Path(tmp) / PROBE.name
            altered.write_text(
                PROBE.read_text(encoding="utf-8").replace(
                    "const PARAGRAPHS: usize = 1_000;",
                    "const PARAGRAPHS: usize = 999;",
                    1,
                ),
                encoding="utf-8",
            )
            with self.assertRaises(provenance.ProvenanceError):
                provenance.fixture_descriptor(altered)

    def test_changed_payload_marker_or_extraction_is_rejected_by_probe_pin(self):
        source = PROBE.read_text(encoding="utf-8")
        changes = (
            ("Replaced text", "Changed payload"),
            ('"/body/p[1]"', '"/body/p[2]"'),
            ("into_file()", "file()"),
        )
        for old, new in changes:
            with self.subTest(old=old):
                with tempfile.TemporaryDirectory() as tmp:
                    altered = Path(tmp) / PROBE.name
                    altered.write_text(source.replace(old, new, 1), encoding="utf-8")
                    with self.assertRaises(provenance.ProvenanceError):
                        provenance.fixture_descriptor(altered)


class DependencyProjection(unittest.TestCase):
    @staticmethod
    def reviewed_probe(root: Path) -> Path:
        probe = root / provenance.PROBE_RELATIVE
        probe.parent.mkdir(parents=True, exist_ok=True)
        probe.write_bytes(PROBE.read_bytes())
        return probe

    def projection_args(self, root: Path, lock: Path, fixture: Path, probe: Path) -> argparse.Namespace:
        return argparse.Namespace(
            workspace=root,
            package_manifest=root / "crates/litchi-ops/Cargo.toml",
            lock=lock,
            toolchain=root / "rust-toolchain.toml",
            probe=probe,
            fixture=fixture,
            features=["docx,xlsx"],
            all_features=False,
            target="x86_64-unknown-linux-gnu",
            profile="release",
        )

    def test_projection_contains_real_lock_graph_and_workload_hashes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text(
                "[workspace]\nmembers = [\"crates/*\"]\n[workspace.dependencies]\nserde = \"1\"\n",
                encoding="utf-8",
            )
            (root / "rust-toolchain.toml").write_text(
                "[toolchain]\nchannel = \"1.95.0\"\ncomponents = [\"clippy\"]\n",
                encoding="utf-8",
            )
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n[features]\ndocx = []\nxlsx = []\n[dependencies]\nserde = { workspace = true }\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\ndependencies = [\"serde\"]\n\n[[package]]\nname = \"serde\"\nversion = \"1.0.0\"\nsource = \"registry+https://example.invalid\"\nchecksum = \"" + "0" * 64 + "\"\n",
                encoding="utf-8",
            )
            fixture = root / "fixture.json"
            probe = self.reviewed_probe(root)
            fixture.write_bytes(provenance.canonical(provenance.fixture_descriptor(probe)))
            projection = provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))
            self.assertEqual(projection["lock"]["packages"][0]["name"], "litchi-ops")
            self.assertEqual(projection["features"]["selected"], ["docx", "xlsx"])
            self.assertEqual(projection["workload"]["route_count"], 109)

    def test_projection_rejects_a_different_workspace_package(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text("[workspace]\n", encoding="utf-8")
            (root / "rust-toolchain.toml").write_text(
                "[toolchain]\nchannel = \"1.95.0\"\n",
                encoding="utf-8",
            )
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-core\"\nversion = \"0.0.1\"\n"
                "[features]\ndocx = []\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-core\"\nversion = \"0.0.1\"\n",
                encoding="utf-8",
            )
            probe = self.reviewed_probe(root)
            fixture = root / "fixture.json"
            fixture.write_bytes(provenance.canonical(provenance.fixture_descriptor(probe)))
            with self.assertRaises(provenance.ProvenanceError):
                provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))

    def test_projection_rejects_fixture_for_another_probe(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text("[workspace]\n", encoding="utf-8")
            (root / "rust-toolchain.toml").write_text("[toolchain]\nchannel = \"1.95.0\"\n", encoding="utf-8")
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n[features]\ndocx = []\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n",
                encoding="utf-8",
            )
            probe = self.reviewed_probe(root)
            fixture = root / "fixture.json"
            fixture.write_bytes(provenance.canonical(provenance.fixture_descriptor(probe)))
            altered = json.loads(fixture.read_text(encoding="utf-8"))
            altered["route_count"] = 108
            fixture.write_bytes(provenance.canonical(altered))
            with self.assertRaises(provenance.ProvenanceError):
                provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))

    def test_projection_rejects_schema_and_probe_hash_without_complete_recipe(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package_dir = root / "crates/litchi-ops"
            package_dir.mkdir(parents=True)
            (root / "Cargo.toml").write_text("[workspace]\n", encoding="utf-8")
            (root / "rust-toolchain.toml").write_text("[toolchain]\nchannel = \"1.95.0\"\n", encoding="utf-8")
            (package_dir / "Cargo.toml").write_text(
                "[package]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n[features]\ndocx = []\n",
                encoding="utf-8",
            )
            lock = root / "Cargo.lock"
            lock.write_text(
                "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n",
                encoding="utf-8",
            )
            fixture = root / "fixture.json"
            probe = self.reviewed_probe(root)
            fixture.write_bytes(
                provenance.canonical(
                    {
                        "schema": "litchi_w8_fixture_descriptor.v1",
                        "probe": {"sha256": provenance.REVIEWED_PROBE_SHA256},
                    }
                )
            )
            with self.assertRaises(provenance.ProvenanceError):
                provenance.dependency_projection(self.projection_args(root, lock, fixture, probe))


class EndpointManifest(unittest.TestCase):
    commit = "a" * 40
    overlay_base = "b" * 40

    def inputs(self, root: Path) -> tuple[argparse.Namespace, dict[str, Path], dict[str, object]]:
        package_dir = root / "crates/litchi-ops"
        package_dir.mkdir(parents=True)
        workspace_manifest = root / "Cargo.toml"
        workspace_manifest.write_text(
            "[workspace]\n[workspace.package]\nversion = \"0.0.1\"\n",
            encoding="utf-8",
        )
        toolchain = root / "rust-toolchain.toml"
        toolchain.write_text(
            "[toolchain]\nchannel = \"1.95.0\"\ncomponents = [\"clippy\"]\n",
            encoding="utf-8",
        )
        package_manifest = package_dir / "Cargo.toml"
        package_manifest.write_text(
            "[package]\nname = \"litchi-ops\"\nversion.workspace = true\n"
            "[features]\ndocx = []\nxlsx = []\n",
            encoding="utf-8",
        )
        probe = root / provenance.PROBE_RELATIVE
        probe.parent.mkdir(parents=True, exist_ok=True)
        probe.write_bytes(PROBE.read_bytes())
        lock = root / "Cargo.lock"
        lock.write_text(
            "version = 4\n\n[[package]]\nname = \"litchi-ops\"\nversion = \"0.0.1\"\n",
            encoding="utf-8",
        )
        fixture = root / "fixture.json"
        fixture_value = provenance.fixture_descriptor(probe)
        fixture.write_bytes(provenance.canonical(fixture_value))

        projection_args = argparse.Namespace(
            workspace=root,
            package_manifest=package_manifest,
            lock=lock,
            toolchain=toolchain,
            probe=probe,
            fixture=fixture,
            features=["docx,xlsx"],
            all_features=False,
            target="x86_64-unknown-linux-gnu",
            profile="release",
        )
        projection_value = provenance.dependency_projection(projection_args)
        projection = root / "projection.json"
        projection.write_bytes(provenance.canonical(projection_value))

        source = root / "source.tar"
        source.write_bytes(b"audited source archive")
        binary = root / "endpoint"
        binary.write_bytes(b"not a native executable in this unit test\n")
        os.chmod(binary, 0o755)
        overlay_patch = root / "compatibility.patch"
        overlay_patch.write_bytes(b"reviewed compatibility overlay")
        source_assembly = root / "source-assembly.json"
        source_assembly_value = provenance.source_assembly(
            argparse.Namespace(
                source=source,
                source_commit=self.commit,
                overlay_kind="compatibility",
                overlay_base_commit=self.overlay_base,
                overlay_patch=overlay_patch,
                description="audited compatibility overlay for the historical baseline",
            )
        )
        source_assembly.write_bytes(provenance.canonical(source_assembly_value))

        compiler = root / "compiler-artifact"
        compiler.write_bytes(b"compiler receipt artifact")
        raw_routes = root / "routes.jsonl"
        raw_routes.write_text("reviewed route receipt\n", encoding="utf-8")
        routes = provenance.fixture_routes(fixture_value)
        validation = root / "validation.json"
        validation_value = {
            "schema": "litchi_file_untimed_validation.v1",
            "source_commit": self.commit,
            "route_count": len(routes),
            "artifacts": {
                "source": provenance.artifact(source),
                "binary": provenance.artifact(binary),
                "lock": provenance.artifact(lock),
                "probe": provenance.artifact(probe),
                "routes": provenance.artifact(raw_routes),
            },
            "outcomes": [
                {
                    "index": index,
                    "route": route,
                    "semantic_equal": True,
                    "semantic_state_valid": True,
                    "source_state_preserved": True,
                    "source_state_exact": True,
                    "exact_noop": True,
                }
                for index, route in enumerate(routes)
            ],
        }
        validation.write_bytes(provenance.canonical(validation_value))
        build = root / "build.json"
        build_value = {
            "schema": "litchi_w8_build_receipt.v1",
            "status": "complete",
            "source_commit": self.commit,
            "source_archive": provenance.artifact(source),
            "source_assembly": provenance.artifact(source_assembly),
            "binary": provenance.artifact(binary),
            "lock": provenance.artifact(lock),
            "probe": provenance.artifact(probe),
            "fixture": provenance.artifact(fixture),
            "dependency_projection_sha256": provenance.digest(projection),
            "workspace_manifest_sha256": provenance.digest(workspace_manifest),
            "package_manifest_sha256": provenance.digest(package_manifest),
            "toolchain_sha256": provenance.digest(toolchain),
            "features": projection_value["features"],
            "target": projection_value["build"]["target"],
            "profile": projection_value["build"]["profile"],
            "build_command": "recorded external release build",
            "compiler_artifacts": [{"name": "rustc", **provenance.artifact(compiler)}],
        }
        build.write_bytes(provenance.canonical(build_value))
        args = argparse.Namespace(
            binary=binary,
            source=source,
            lock=lock,
            probe=probe,
            fixture=fixture,
            dependency_projection=projection,
            workspace=root,
            workspace_manifest=workspace_manifest,
            package_manifest=package_manifest,
            toolchain=toolchain,
            source_assembly=source_assembly,
            build_receipt=build,
            validation_receipt=validation,
            source_commit=self.commit,
            features=["docx,xlsx"],
            all_features=False,
            target="x86_64-unknown-linux-gnu",
            profile="release",
        )
        paths = {"lock": lock, "validation": validation, "projection": projection}
        values = {"routes": routes, "projection": projection_value}
        return args, paths, values

    def test_build_receipt_generator_binds_exact_release_command_without_cargo(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            args, paths, _values = self.inputs(root)
            cargo = root / "cargo-shim"
            compiler = root / "rustc"
            cargo.write_bytes(b"reviewed cargo shim")
            compiler.write_bytes(b"reviewed rustc")
            os.chmod(cargo, 0o755)
            os.chmod(compiler, 0o755)
            receipt = root / "generated-build.json"
            command = (
                f"CARGO_BUILD_JOBS=1 CARGO_TARGET_DIR={root / 'target'} {cargo} build "
                "--offline --locked --target x86_64-unknown-linux-gnu --release "
                "-p litchi-ops --all-features --example file_w8_qualified_probe"
            )
            build_args = argparse.Namespace(
                source=args.source,
                source_assembly=args.source_assembly,
                binary=args.binary,
                lock=args.lock,
                probe=args.probe,
                fixture=args.fixture,
                dependency_projection=paths["projection"],
                workspace=root,
                workspace_manifest=args.workspace_manifest,
                package_manifest=args.package_manifest,
                toolchain=args.toolchain,
                cargo=cargo,
                output=receipt,
                source_commit=self.commit,
                compiler_artifacts=[f"rustc={compiler}"],
                build_command=command,
                features=["docx", "xlsx"],
                all_features=False,
                target="x86_64-unknown-linux-gnu",
                profile="release",
            )
            row = provenance.write_build_receipt(build_args)
            value = json.loads(receipt.read_text(encoding="utf-8"))
            self.assertEqual(row["sha256"], provenance.digest(receipt))
            self.assertEqual(value["schema"], "litchi_w8_build_receipt.v1")
            self.assertEqual(value["cargo"]["sha256"], provenance.digest(cargo))
            self.assertEqual(value["compiler_artifacts"][0]["name"], "rustc")
            self.assertIn("--target x86_64-unknown-linux-gnu", value["build_command"])
            self.assertIn("--release", value["build_command"])

            bad = argparse.Namespace(**vars(build_args))
            bad.output = root / "bad-build.json"
            bad.build_command = f"{cargo} build"
            with self.assertRaises(provenance.ProvenanceError):
                provenance.write_build_receipt(bad)
            self.assertFalse((root / "bad-build.json").exists())

    @staticmethod
    def runner_module():
        path = ROOT / "docs/unified-ops/results/0362h/run.py"
        spec = importlib.util.spec_from_file_location("w8_file_run", path)
        if spec is None or spec.loader is None:
            raise AssertionError(f"cannot load {path}")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def test_endpoint_binds_projection_and_receipts_to_current_inputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            args, paths, values = self.inputs(Path(tmp))
            endpoint = provenance.endpoint_manifest(args)
            self.assertEqual(endpoint["source_commit"], self.commit)
            self.assertEqual(endpoint["dependency_projection"]["sha256"], provenance.digest(paths["projection"]))
            self.assertEqual(len(values["routes"]), 109)

            paths["lock"].write_text(
                paths["lock"].read_text(encoding="utf-8") + "\n# stale join\n",
                encoding="utf-8",
            )
            with self.assertRaises(provenance.ProvenanceError):
                provenance.endpoint_manifest(args)

    def test_endpoint_rejects_source_archive_and_overlay_patch_mutations(self):
        with tempfile.TemporaryDirectory() as tmp:
            args, _paths, _values = self.inputs(Path(tmp))
            provenance.endpoint_manifest(args)
            args.source.write_bytes(b"mutated source archive")
            with self.assertRaises(provenance.ProvenanceError):
                provenance.endpoint_manifest(args)

            args.source.write_bytes(b"audited source archive")
            args.overlay_patch = Path(tmp) / "compatibility.patch"
            args.overlay_patch.write_bytes(b"mutated adapter patch")
            with self.assertRaises(provenance.ProvenanceError):
                provenance.endpoint_manifest(args)

    def test_endpoint_rejects_short_source_commit_and_incomplete_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            args, paths, _ = self.inputs(Path(tmp))
            args.source_commit = "f2bbbc8d07"
            with self.assertRaises(provenance.ProvenanceError):
                provenance.endpoint_manifest(args)

            args.source_commit = self.commit
            value = json.loads(paths["validation"].read_text(encoding="utf-8"))
            value["route_count"] = 108
            paths["validation"].write_bytes(provenance.canonical(value))
            with self.assertRaises(provenance.ProvenanceError):
                provenance.endpoint_manifest(args)

    def test_runtime_rejects_projection_with_only_schema_and_probe_hash(self):
        with tempfile.TemporaryDirectory() as tmp:
            args, paths, _ = self.inputs(Path(tmp))
            endpoint = provenance.endpoint_manifest(args)
            endpoint_path = Path(tmp) / "endpoint.json"
            endpoint_path.write_bytes(provenance.canonical(endpoint))
            runner = self.runner_module()
            checked = runner.endpoint_manifest(endpoint_path, args.binary, "test")
            self.assertEqual(len(checked["routes"]), 109)

            paths["projection"].write_bytes(
                provenance.canonical(
                    {
                        "schema": "litchi_w8_dependency_projection.v1",
                        "workload": {"probe_sha256": provenance.REVIEWED_PROBE_SHA256},
                    }
                )
            )
            endpoint["dependency_projection"] = provenance.artifact(paths["projection"])
            endpoint["dependency_projection_sha256"] = provenance.digest(paths["projection"])
            endpoint_path.write_bytes(provenance.canonical(endpoint))
            with self.assertRaises(RuntimeError):
                runner.endpoint_manifest(endpoint_path, args.binary, "partial")

    def test_runtime_rejects_stale_workspace_package_version(self):
        with tempfile.TemporaryDirectory() as tmp:
            args, paths, _ = self.inputs(Path(tmp))
            endpoint = provenance.endpoint_manifest(args)
            endpoint_path = Path(tmp) / "endpoint.json"
            endpoint_path.write_bytes(provenance.canonical(endpoint))
            runner = self.runner_module()

            projection = json.loads(paths["projection"].read_text(encoding="utf-8"))
            projection["workspace"]["package_version"] = "9.9.9"
            paths["projection"].write_bytes(provenance.canonical(projection))
            endpoint["dependency_projection"] = provenance.artifact(paths["projection"])
            endpoint["dependency_projection_sha256"] = provenance.digest(paths["projection"])
            endpoint_path.write_bytes(provenance.canonical(endpoint))
            with self.assertRaisesRegex(RuntimeError, "projection package version differs from the manifests"):
                runner.endpoint_manifest(endpoint_path, args.binary, "stale-version")


if __name__ == "__main__":
    unittest.main()
