from __future__ import annotations

import copy
import json
import hashlib
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

from tools import raw_api_coverage as coverage


def rustdoc_fixture() -> dict:
    """One raw callable and one explicitly semantic callable."""
    return {
        "root": 0,
        "includes_private": False,
        "index": {
            "0": {
                "name": "fixture",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"module": {"items": [1, 2, 3], "is_crate": True}},
            },
            "1": {
                "name": "SourceLimits",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"struct": {"fields": [], "impls": []}},
            },
            "2": {
                "name": "parse",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"function": {
                    "sig": {"inputs": [
                        ["data", {"primitive": "unit"}],
                        ["limits", {"borrowed_ref": {
                            "is_mutable": False,
                            "type": {"resolved_path": {"id": 1, "args": None}},
                        }}],
                    ], "output": {"primitive": "unit"}},
                    "generics": {"params": []},
                }},
            },
            "3": {
                "name": "describe",
                "crate_id": 0,
                "visibility": "public",
                "inner": {"function": {
                    "sig": {"inputs": [], "output": {"primitive": "unit"}},
                    "generics": {"params": []},
                }},
            },
        },
        "paths": {
            "0": {"path": ["fixture"], "kind": "module", "crate_id": 0},
            "1": {"path": ["fixture", "SourceLimits"], "kind": "struct", "crate_id": 0},
            "2": {"path": ["fixture", "parse"], "kind": "function", "crate_id": 0},
            "3": {"path": ["fixture", "describe"], "kind": "function", "crate_id": 0},
        },
    }


def glossary_fixture(path: Path) -> None:
    path.write_text(json.dumps({
        "profiles": [{
            "id": "fixture-profile",
            "package": "fixture",
            "default_features": True,
            "features": [],
        }],
    }), encoding="utf-8")


def entry(path: str, kind: str, disposition: str, *, signature=None) -> dict:
    value = {
        "profile": "fixture-profile",
        "path": path,
        "kind": kind,
        "disposition": disposition,
        "source": {
            "path": "crates/fixture/src/lib.rs",
            "blob": "a" * 40,
            "line": 1,
            "symbol": path.rsplit("::", 1)[-1],
            "symbol_kind": "rust_item",
        },
        "evidence": "synthetic rustdoc fixture evidence",
    }
    if signature is not None:
        value["required_signature"] = signature
    return value


def manifest(entries: list[dict], *, state: str = "complete") -> dict:
    value = {
        "schema_version": 1,
        "source_commit": "b" * 40,
        "coverage": {
            "state": state,
            "reason": "synthetic coverage fixture",
            "expected_profile_source": "tools/api_glossary.json",
            "expected_profile_source_blob": "c" * 40,
            "expected_workspace_source": "Cargo.toml",
            "expected_workspace_source_blob": "d" * 40,
            "approved_cargo_lock_sha256": "e" * 64,
            "capture_policy": {
                "runner_path": "/opt/f5/run_safe_gate.py",
                "runner_sha256": "f" * 64,
                "cap_wrapper_path": "/opt/f5/bin/cargo",
                "cap_wrapper_sha256": "a" * 64,
                "underlying_cargo_path": "/opt/f5/coord/cargo",
                "underlying_cargo_sha256": "b" * 64,
                "environment": dict(coverage.CAPTURE_POLICY_ENV),
                "guards": dict(coverage.CAPTURE_POLICY_GUARDS),
            },
        },
        "profiles": [{
            "id": "fixture-profile",
            "package": "fixture",
            "default_features": True,
            "features": [],
        }],
        "entries": entries,
        "lanes": [],
    }
    for item in entries:
        if item["disposition"] != "active_lane":
            continue
        value["lanes"].append({
            "id": item["lane"],
            "package": "fixture",
            "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "crates/fixture/src/lib.rs",
                "blob": "a" * 40,
                "line": 1,
                "symbol": "lane-anchor",
                "symbol_kind": "descriptive",
            },
            "evidence": "synthetic active lane",
        })
    return value


class RawApiCoverageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.capture_dirs: list[Path] = []
        self.glossary = self.root / "glossary.json"
        glossary_fixture(self.glossary)

    def tearDown(self) -> None:
        for path in self.capture_dirs:
            shutil.rmtree(path, ignore_errors=True)
        self.tmp.cleanup()

    @staticmethod
    def sha256(path: Path) -> str:
        return hashlib.sha256(path.read_bytes()).hexdigest()

    def capture_fixture(self) -> tuple[Path, coverage.Profile, dict, Path, Path, dict]:
        external = Path(tempfile.mkdtemp(prefix="raw-api-receipt-", dir=Path.cwd().parent))
        self.capture_dirs.append(external)
        root = external / "source"
        root.mkdir()
        (root / ".gitignore").write_text(
            "*.lock\ntarget/\nvendor/\nmanifest.json\nfixtures/\n", encoding="utf-8"
        )
        (root / "Cargo.toml").write_text("[workspace]\n", encoding="utf-8")
        (root / "Cargo.lock").write_text("approved lock\n", encoding="utf-8")
        (root / "rust-toolchain.toml").write_text(
            "[toolchain]\nchannel = \"1.95.0\"\n", encoding="utf-8"
        )
        subprocess.run(["git", "init", "-q"], cwd=root, check=True)
        subprocess.run(["git", "config", "user.email", "fixture@example.invalid"], cwd=root, check=True)
        subprocess.run(["git", "config", "user.name", "fixture"], cwd=root, check=True)
        subprocess.run(["git", "add", ".gitignore", "Cargo.toml", "rust-toolchain.toml"], cwd=root, check=True)
        subprocess.run(["git", "commit", "-qm", "fixture source"], cwd=root, check=True)
        commit = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=root, text=True
        ).strip()
        tree = coverage._git_tree(root, commit)
        self.assertIsNotNone(tree)

        toolchain = external / "toolchain"
        toolchain.mkdir()
        rustc = toolchain / "rustc"
        rustc.write_text("rustc fixture\n", encoding="utf-8")
        runner = external / "runner" / "run_safe_gate.py"
        runner.parent.mkdir()
        runner.write_text("runner fixture\n", encoding="utf-8")
        cap_wrapper = external / "runner" / "bin" / "cargo"
        cap_wrapper.parent.mkdir()
        cap_wrapper.write_text(
            "#!/bin/sh\n"
            "LITCHI_GATE_TARGET_TRACE=requested_target\n"
            "printf '%s\\t%s\\t%s\\n' \"$$\" \"$command\" \"$requested_target\"\n",
            encoding="utf-8",
        )
        cap_wrapper.chmod(0o755)
        underlying_cargo = external / "runner" / "coord" / "cargo"
        underlying_cargo.parent.mkdir()
        underlying_cargo.write_text(
            "#!/bin/sh\nsystemd-run MemoryMax=${LITCHI_MEM_MAX:-20G}\n", encoding="utf-8"
        )
        underlying_cargo.chmod(0o755)
        target = external / "target"
        target.mkdir()
        tmpdir = external / "tmp"
        tmpdir.mkdir()
        artifact = external / "capture" / "fixture-profile.json"
        artifact.parent.mkdir()
        artifact.write_text(json.dumps(rustdoc_fixture()), encoding="utf-8")
        trace = external / "capture" / "fixture-profile.trace"
        trace.write_text(f"123\trustdoc\t{target}\n", encoding="utf-8")
        receipt_path = external / "receipts" / "fixture-profile.json"
        receipt_path.parent.mkdir()
        profile = coverage.Profile("fixture-profile", "fixture", False, ("feature-a",))
        lock_sha = self.sha256(root / "Cargo.lock")
        policy = {
            "runner_path": str(runner),
            "runner_sha256": self.sha256(runner),
            "cap_wrapper_path": str(cap_wrapper),
            "cap_wrapper_sha256": self.sha256(cap_wrapper),
            "underlying_cargo_path": str(underlying_cargo),
            "underlying_cargo_sha256": self.sha256(underlying_cargo),
            "environment": dict(coverage.CAPTURE_POLICY_ENV),
            "guards": dict(coverage.CAPTURE_POLICY_GUARDS),
        }
        capture_env = dict(policy["environment"])
        capture_env.update({
            "CARGO_TARGET_DIR": str(target),
            "LITCHI_GATE_TARGET_ROOT": str(target),
            "LITCHI_GATE_TARGET_TRACE": str(trace),
            "TMPDIR": str(tmpdir),
        })
        receipt = {
            "schema": coverage.CAPTURE_RECEIPT_SCHEMA,
            "profile": profile.profile_id,
            "package": profile.package,
            "default_features": profile.default_features,
            "features": list(profile.features),
            "argv": [
                str(cap_wrapper), "rustdoc", "--package", "fixture", "--lib",
                "--target-dir", str(target), "--no-default-features", "--features", "feature-a",
                "--offline", "--locked", "--", "-Zunstable-options", "--output-format", "json",
                "--document-hidden-items",
            ],
            "source_commit": commit,
            "source_tree": tree,
            "source_root": str(root),
            "target_dir": str(target),
            "cargo_lock_path": "Cargo.lock",
            "cargo_lock_sha256": lock_sha,
            "rust_toolchain_path": "rust-toolchain.toml",
            "rust_toolchain_sha256": self.sha256(root / "rust-toolchain.toml"),
            "cargo_path": str(cap_wrapper),
            "cargo_sha256": self.sha256(cap_wrapper),
            "rustc_path": str(rustc),
            "rustc_sha256": self.sha256(rustc),
            "runner_path": str(runner),
            "runner_sha256": self.sha256(runner),
            "cap_wrapper_path": str(cap_wrapper),
            "cap_wrapper_sha256": self.sha256(cap_wrapper),
            "underlying_cargo_path": str(underlying_cargo),
            "underlying_cargo_sha256": self.sha256(underlying_cargo),
            "capture_env": capture_env,
            "guards": dict(coverage.CAPTURE_POLICY_GUARDS),
            "resource": {
                "target_pre_bytes": 0,
                "target_live_peak_bytes": 1,
                "target_post_bytes": 1,
                "peak_rss_bytes": 1,
                "free_bytes_after": coverage.CAPTURE_POLICY_GUARDS["free_min_bytes"] + 1,
                "target_live_guard_triggered": False,
            },
            "execution_trace": {
                "path": str(trace),
                "bytes": trace.stat().st_size,
                "sha256": self.sha256(trace),
                "rows": [{"pid": "123", "command": "rustdoc", "target": str(target)}],
            },
            "artifact": {
                "path": str(artifact),
                "bytes": artifact.stat().st_size,
                "sha256": self.sha256(artifact),
            },
        }
        receipt_path.write_text(json.dumps(receipt), encoding="utf-8")
        return root, profile, receipt, artifact, receipt_path, policy

    def loaded(self, value: dict, *, require_complete: bool = False,
               extra_rustdoc: bool = False) -> dict:
        manifest_path = self.root / "manifest.json"
        manifest_path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(manifest_path)
        rustdoc = self.root / "fixture.json"
        rustdoc.write_text(json.dumps(rustdoc_fixture()), encoding="utf-8")
        bindings = {"fixture-profile": rustdoc}
        if extra_rustdoc:
            bindings["extra-profile"] = rustdoc
        return coverage.validate_manifest(
            loaded, root=self.root, glossary_path=self.glossary,
            rustdoc=bindings, require_complete=require_complete,
            verify_sources=False,
        )

    @staticmethod
    def raw_signature(limits_type: str = "fixture::SourceLimits") -> dict:
        return {
            "inputs": ["data", "limits"],
            "limits_argument": "limits",
            "limits_type": limits_type,
        }

    def test_capture_receipt_binds_profile_source_tools_and_artifact(self) -> None:
        root, profile, receipt, artifact, receipt_path, policy = self.capture_fixture()
        manifest_value = {
            "source_commit": receipt["source_commit"],
            "coverage": {
                "approved_cargo_lock_sha256": receipt["cargo_lock_sha256"],
                "capture_policy": policy,
            },
        }
        parsed = coverage._capture_receipt(receipt, "receipt[fixture-profile]")
        overlays = coverage._source_overlay_failures(root, receipt["cargo_lock_sha256"])
        self.assertEqual(overlays, [])
        self.assertEqual(
            coverage._validate_capture_receipt(
                manifest_value, profile, artifact, receipt_path, root, parsed, overlays
            ),
            [],
        )

        wrong_features = dict(parsed)
        wrong_features["features"] = []
        self.assertTrue(
            any("features differ" in failure for failure in coverage._validate_capture_receipt(
                manifest_value, profile, artifact, receipt_path, root,
                wrong_features, overlays
            ))
        )

        wrong_artifact = dict(parsed)
        wrong_artifact["artifact"] = dict(parsed["artifact"])
        wrong_artifact["artifact"]["bytes"] += 1
        self.assertTrue(
            any("artifact bytes/SHA" in failure for failure in coverage._validate_capture_receipt(
                manifest_value, profile, artifact, receipt_path, root,
                wrong_artifact, overlays
            ))
        )

        stale = dict(parsed)
        stale["source_commit"] = "0" * 40
        self.assertTrue(
            any("source_commit is stale" in failure for failure in coverage._validate_capture_receipt(
                manifest_value, profile, artifact, receipt_path, root,
                stale, overlays
            ))
        )

        wrong_seed = {
            "source_commit": receipt["source_commit"],
            "coverage": {
                "approved_cargo_lock_sha256": "0" * 64,
                "capture_policy": policy,
            },
        }
        self.assertTrue(
            any("approved seeded lock" in failure for failure in coverage._validate_capture_receipt(
                wrong_seed, profile, artifact, receipt_path, root, parsed, overlays
            ))
        )

    def test_capture_receipt_rejects_source_overlays_and_wrong_seed_lock(self) -> None:
        root, profile, receipt, artifact, receipt_path, _ = self.capture_fixture()
        del profile, artifact, receipt_path
        (root / "vendor" / "overlay").mkdir(parents=True)
        (root / "vendor" / "overlay" / "fixture.bin").write_bytes(b"overlay")
        (root / "manifest.json").write_text("overlay", encoding="utf-8")
        (root / "fixtures").mkdir()
        (root / "fixtures" / "fixture.json").write_text("overlay", encoding="utf-8")
        failures = coverage._source_overlay_failures(root, receipt["cargo_lock_sha256"])
        self.assertTrue(any("vendor/overlay" in failure for failure in failures), failures)
        self.assertTrue(any("manifest.json" in failure for failure in failures), failures)
        self.assertTrue(any("fixtures/fixture.json" in failure for failure in failures), failures)
        wrong_lock = coverage._source_overlay_failures(root, "0" * 64)
        self.assertTrue(any("Cargo.lock" in failure for failure in wrong_lock), wrong_lock)

    def test_capture_receipt_rejects_lock_symlink(self) -> None:
        root, _, receipt, _, _, _ = self.capture_fixture()
        target = root.parent / "alternate-lock"
        target.write_text("approved lock\n", encoding="utf-8")
        (root / "Cargo.lock").unlink()
        (root / "Cargo.lock").symlink_to(target)
        failures = coverage._source_overlay_failures(root, receipt["cargo_lock_sha256"])
        self.assertTrue(any("Cargo.lock" in failure for failure in failures), failures)
        self.assertFalse(coverage._regular_file_with_components(root / "Cargo.lock"))

    def test_capture_receipt_rejects_runner_environment_and_trace_tampering(self) -> None:
        root, profile, receipt, artifact, receipt_path, policy = self.capture_fixture()
        manifest_value = {
            "source_commit": receipt["source_commit"],
            "coverage": {
                "approved_cargo_lock_sha256": receipt["cargo_lock_sha256"],
                "capture_policy": policy,
            },
        }
        overlays = coverage._source_overlay_failures(root, receipt["cargo_lock_sha256"])
        bad_env = dict(receipt)
        bad_env["capture_env"] = dict(receipt["capture_env"])
        bad_env["capture_env"]["CARGO_NET_OFFLINE"] = "false"
        parsed = coverage._capture_receipt(bad_env, "receipt[fixture-profile]")
        self.assertTrue(any("capture environment" in failure for failure in coverage._validate_capture_receipt(
            manifest_value, profile, artifact, receipt_path, root, parsed, overlays
        )))

        bad_runner = dict(receipt)
        bad_runner["runner_sha256"] = "0" * 64
        parsed = coverage._capture_receipt(bad_runner, "receipt[fixture-profile]")
        self.assertTrue(any("runner_sha256" in failure for failure in coverage._validate_capture_receipt(
            manifest_value, profile, artifact, receipt_path, root, parsed, overlays
        )))

        Path(receipt["execution_trace"]["path"]).write_text(
            f"123\tcheck\t{receipt['target_dir']}\n", encoding="utf-8"
        )
        parsed = coverage._capture_receipt(receipt, "receipt[fixture-profile]")
        failures = coverage._validate_capture_receipt(
            manifest_value, profile, artifact, receipt_path, root, parsed, overlays
        )
        self.assertTrue(any("execution trace" in failure for failure in failures), failures)

    def test_receipt_argument_rejects_duplicate_profiles(self) -> None:
        with self.assertRaises(coverage.CoverageError):
            coverage.parse_receipt_args(["fixture-profile=a", "fixture-profile=b"])

    def test_direct_cli_does_not_create_bytecode_before_overlay_scan(self) -> None:
        source = Path(self.tmp.name) / "cli-source"
        tools = source / "tools"
        tools.mkdir(parents=True)
        for name in ("raw_api_coverage.py", "check_api_glossary.py", "rustdoc_return_types.py",
                     "raw_api_coverage.json", "api_glossary.json"):
            shutil.copyfile(Path(__file__).with_name(name), tools / name)
        environment = os.environ.copy()
        environment.pop("PYTHONDONTWRITEBYTECODE", None)
        result = subprocess.run(
            [
                sys.executable, os.fspath(tools / "raw_api_coverage.py"),
                "--root", os.fspath(source), "--manifest", os.fspath(tools / "raw_api_coverage.json"),
                "--glossary", os.fspath(tools / "api_glossary.json"), "--json",
            ], cwd=source, env=environment, capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse((tools / "__pycache__").exists())
        self.assertEqual(coverage.CAPTURE_POLICY_ENV["PYTHONDONTWRITEBYTECODE"], "1")

    def test_exact_census_accepts_registered_raw_signature_and_semantic_anchor(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value, require_complete=True)
        self.assertTrue(report["complete"], report)
        self.assertEqual(report["callable_counts"], {"fixture-profile": 2})
        self.assertEqual(report["unclassified"], [])

    def test_missing_callable_is_unclassified_even_when_name_is_not_raw(self) -> None:
        value = manifest([entry(
            "fixture::parse", "raw_function", "required_raw_limits",
            signature=self.raw_signature(),
        )])
        report = self.loaded(value)
        self.assertEqual(report["unclassified"], ["fixture-profile::fixture::describe"])
        self.assertFalse(report["complete"])

    def test_raw_signature_type_mismatch_fails_closed(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature("fixture::WrongLimits")),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value)
        self.assertEqual(len(report["signature_failures"]), 1)
        self.assertIn("limits type identity", report["signature_failures"][0])

    def test_raw_signature_accepts_external_mce_reexport_by_nominal_identity(self) -> None:
        document = rustdoc_fixture()
        method_id = "mce-limit-method"
        defining_id = "mce-limit-defining-item"
        document["index"]["0"]["inner"]["module"]["items"].append(method_id)
        document["index"][method_id] = {
            "name": "parse_mce", "crate_id": 0, "visibility": "public",
            "inner": {"function": {
                "sig": {"inputs": [
                    ["data", {"primitive": "unit"}],
                    ["limits", {"borrowed_ref": {
                        "is_mutable": False,
                        "type": {"resolved_path": {"id": defining_id, "args": None}},
                    }}],
                ], "output": {"primitive": "unit"}},
                "generics": {"params": []},
            }},
        }
        document["paths"][method_id] = {
            "path": ["fixture", "parse_mce"], "kind": "function", "crate_id": 0,
        }
        document["paths"][defining_id] = {
            "path": ["litchi_ooxml_common", "mce", "model", "Limits"],
            "kind": "struct", "crate_id": 81,
        }
        signature = self.raw_signature("litchi_ooxml_common::mce::Limits")
        entry_value = {"path": "fixture::parse_mce", "required_signature": signature}
        surface = coverage.glossary_gate.Surface(document, "fixture")
        self.assertIsNone(coverage.check_required_signature(surface, entry_value))

        wrong = copy.deepcopy(document)
        wrong_id = "mce-limit-unrelated-item"
        wrong["paths"][wrong_id] = {
            "path": ["litchi_ooxml_common", "mce", "other", "Limits"],
            "kind": "struct", "crate_id": 82,
        }
        wrong["index"][method_id]["inner"]["function"]["sig"]["inputs"][1][1][
            "borrowed_ref"
        ]["type"]["resolved_path"]["id"] = wrong_id
        surface = coverage.glossary_gate.Surface(wrong, "fixture")
        self.assertIn("limits type identity", coverage.check_required_signature(surface, entry_value))

        ambiguous = copy.deepcopy(document)
        ambiguous["paths"]["mce-limit-duplicate"] = {
            "path": ["litchi_ooxml_common", "mce", "model", "Limits"],
            "kind": "struct", "crate_id": 83,
        }
        surface = coverage.glossary_gate.Surface(ambiguous, "fixture")
        self.assertIn("limits type identity", coverage.check_required_signature(surface, entry_value))

    def test_active_lane_and_incomplete_state_fail_final_mode(self) -> None:
        active = entry("fixture::describe", "semantic_method", "active_lane")
        active["lane"] = "0362ul-test"
        value = manifest([active])
        value["lanes"] = []
        value["lanes"].append({
            "id": "0362ul-test", "package": "fixture", "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "crates/fixture/src/lib.rs", "blob": "a" * 40,
                "line": 1, "symbol": "lane-anchor", "symbol_kind": "descriptive",
            },
            "evidence": "synthetic active lane",
        })
        value["coverage"]["state"] = "incomplete"
        report = self.loaded(value, require_complete=True)
        self.assertFalse(report["complete"])
        self.assertEqual(report["active_entries"], ["fixture-profile::fixture::describe"])
        self.assertIn("failure", report)

    def test_private_removed_is_stale_if_rustdoc_still_exports_it(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "removed_callable", "private_removed"),
        ])
        report = self.loaded(value)
        self.assertEqual(report["stale_entries"], [
            "fixture-profile::fixture::describe: private_removed item remains public",
        ])

    def test_manifest_rejects_generic_disposition_and_unknown_fields(self) -> None:
        value = manifest([entry("fixture::parse", "raw_function", "out_of_scope")])
        with self.assertRaises(coverage.CoverageError):
            path = self.root / "manifest.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            coverage.load_manifest(path)

    def test_extra_rustdoc_binding_cannot_make_final_mode_pass(self) -> None:
        value = manifest([
            entry("fixture::parse", "raw_function", "required_raw_limits",
                  signature=self.raw_signature()),
            entry("fixture::describe", "semantic_method", "semantic_value"),
        ])
        report = self.loaded(value, require_complete=True, extra_rustdoc=True)
        self.assertEqual(report["extra_inputs"], ["extra-profile"])
        self.assertFalse(report["complete"])

    def test_undeclared_active_lane_is_rejected(self) -> None:
        active = entry("fixture::describe", "semantic_method", "active_lane")
        active["lane"] = "0362ul-missing"
        value = manifest([active])
        value["lanes"] = []
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaises(coverage.CoverageError):
            coverage.load_manifest(path)

    def test_callable_source_anchor_must_be_rust_item(self) -> None:
        value = manifest([entry("fixture::describe", "semantic_method", "semantic_value")])
        value["entries"][0]["source"]["symbol_kind"] = "descriptive"
        value["entries"][0]["source"]["symbol"] = "arbitrary-label"
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaisesRegex(coverage.CoverageError, "rust_item"):
            coverage.load_manifest(path)

    def test_source_anchor_requires_blob_bounds_and_verified_symbol(self) -> None:
        root = Path(__file__).resolve().parents[1]
        head = coverage._git_head(root)
        self.assertIsNotNone(head)
        blob = coverage._git_blob(root, head, "tools/check_api_glossary.py")
        self.assertIsNotNone(blob)
        value = manifest([])
        value["source_commit"] = head
        value["coverage"]["expected_profile_source_blob"] = coverage._git_blob(
            root, head, "tools/api_glossary.json"
        )
        value["coverage"]["expected_workspace_source_blob"] = coverage._git_blob(
            root, head, "Cargo.toml"
        )
        value["lanes"] = [{
            "id": "test-source", "package": "fixture", "scope": "fixture",
            "disposition": "active_lane",
            "source": {
                "path": "tools/check_api_glossary.py", "blob": blob,
                "line": 1, "symbol": "module", "symbol_kind": "descriptive",
            },
            "evidence": "source identity test",
        }]
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        loaded = coverage.load_manifest(path)
        self.assertEqual(coverage.verify_source_anchors(loaded, root), [])

        tree = coverage._git_object(root, head, "tools")
        self.assertIsNotNone(tree)
        loaded["source_commit"] = tree[0]
        self.assertIn("does not resolve to a Git commit", coverage.verify_source_anchors(loaded, root)[0])
        loaded["source_commit"] = head

        loaded["lanes"][0]["source"]["blob"] = "c" * 40
        self.assertIn("blob", coverage.verify_source_anchors(loaded, root)[0])
        loaded["lanes"][0]["source"]["blob"] = blob
        loaded["lanes"][0]["source"]["line"] = 999999
        self.assertIn("exceeds", coverage.verify_source_anchors(loaded, root)[0])

        loaded["lanes"][0]["source"]["line"] = 1
        loaded["lanes"][0]["source"]["symbol"] = "not_the_function"
        loaded["lanes"][0]["source"]["symbol_kind"] = "rust_item"
        self.assertIn("does not contain rust item", coverage.verify_source_anchors(loaded, root)[0])

        loaded["lanes"][0]["source"]["path"] = "tools"
        loaded["lanes"][0]["source"]["blob"] = tree[0]
        loaded["lanes"][0]["source"]["symbol_kind"] = "descriptive"
        self.assertIn("not a blob", coverage.verify_source_anchors(loaded, root)[0])

    def test_source_anchor_schema_requires_symbol_kind(self) -> None:
        value = manifest([entry("fixture::describe", "semantic_method", "semantic_value")])
        value["entries"][0]["source"].pop("symbol_kind")
        path = self.root / "manifest.json"
        path.write_text(json.dumps(value), encoding="utf-8")
        with self.assertRaises(coverage.CoverageError):
            coverage.load_manifest(path)


if __name__ == "__main__":
    unittest.main()
