#!/usr/bin/env python3
"""Source-only tests for the signed-native v2 consumer receipt boundary."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import os
import subprocess
import tempfile
import unittest
from unittest import mock

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import prepare_signed_native_inputs_v2 as policy  # noqa: E402
from native_qualification import signed_native_v2_consumer as consumer  # noqa: E402


def canonical(value: object) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


class SignedNativeV2ConsumerTests(unittest.TestCase):
    def _git_repo(self, root: Path, files: dict[str, bytes]) -> None:
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
        subprocess.run(["git", "-C", str(root), "config", "user.name", "signed-native test"], check=True)
        for relative, payload in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(payload)
        subprocess.run(["git", "-C", str(root), "add", "."], check=True)
        subprocess.run(["git", "-C", str(root), "commit", "-qm", "fixture"], check=True)

    def _fixture_and_probes(self, source: Path) -> tuple[dict[str, dict[str, object]], dict[str, object]]:
        fixture_relative = "test-data/signed/fixture.bin"
        fixture_payload = b"signed fixture\n"
        (source / fixture_relative).parent.mkdir(parents=True, exist_ok=True)
        (source / fixture_relative).write_bytes(fixture_payload)
        fixture_value = {
            "schema": "litchi_signed_fixture_manifest.v1",
            "fixtures": [
                {
                    "name": "synthetic",
                    "format": "DOCX",
                    "origin": "Opc",
                    "relative_path": fixture_relative,
                    "input_sha256": hashlib.sha256(fixture_payload).hexdigest(),
                    "signed_sha256": hashlib.sha256(fixture_payload).hexdigest(),
                }
            ],
        }
        fixture_path = source / "fixtures.json"
        fixture_path.write_bytes(canonical(fixture_value))
        probes: dict[str, dict[str, object]] = {}
        for index, name in enumerate(("probe", "qualification_driver", "validator", "memory_collector", "final_collector")):
            relative = f"reviewed/{name}.input"
            payload = (
                'const SCHEMA: &str = "litchi_signed_native_probe.v1";\n'
                '#[cfg(feature = "legacy_api")]\n'
                'fn legacy_api_marker() { let _ = "edit_with_limits"; }\n'
                '#[cfg(not(feature = "legacy_api"))]\n'
                'fn current_api_marker() {}\n'
                'fn main() {\n'
                '    let args: Vec<String> = std::env::args().collect();\n'
                '}\n'
                if name == "probe"
                else f"input-{index}\n"
            ).encode()
            path = source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(payload)
            probes[name] = {
                "path": relative,
                "bytes": len(payload),
                "sha256": hashlib.sha256(payload).hexdigest(),
            }
        return probes, {
            "path": "fixtures.json",
            "bytes": fixture_path.stat().st_size,
            "sha256": hashlib.sha256(fixture_path.read_bytes()).hexdigest(),
        }

    def _fixture_patch(self, probes: dict[str, dict[str, object]], fixture: dict[str, object]):
        role_locks = {
            "workspace": {"path": "Cargo.lock", "baseline": {"sha256": "a" * 64, "bytes": 1}, "candidate": {"sha256": "b" * 64, "bytes": 1}},
            "perf": {"path": "perf.lock", "baseline": {"sha256": "c" * 64, "bytes": 1}, "candidate": {"sha256": "d" * 64, "bytes": 1}},
            "native": {"path": "native.lock", "baseline": {"sha256": "e" * 64, "bytes": 1}, "candidate": {"sha256": "f" * 64, "bytes": 1}},
        }
        manifests = ("Cargo.toml", "perf.toml", "native.toml")
        toolchain_paths = ("rust-toolchain.toml", ".cargo/config.toml")
        vendor_roots = ("vendor/allsorts", "vendor/flate2", "vendor/zlib-rs")
        fake_binding = {
            "status": "bound",
            "source_manifests": list(manifests),
            "role_locks": role_locks,
            "vendor_roots": list(vendor_roots),
            "fixture": fixture,
            "probe_inputs": probes,
            "toolchain_paths": list(toolchain_paths),
            "toolchain": {path: {"path": path, "bytes": 1, "sha256": str(index) * 64} for index, path in enumerate(toolchain_paths, 1)},
            "source": {
                "commit": "1" * 40,
                "git_tree": "2" * 40,
                "tree_sha256": "3" * 64,
                "archive": {
                    "path": "/tmp/reviewed-signed-native.tar",
                    "bytes": 1,
                    "sha256": "4" * 64,
                    "format": "git-tree-tar-v1",
                    "prefix": "capture/",
                    "git_tree": "2" * 40,
                    "tree_sha256": "3" * 64,
                    "ignored_files": [],
                },
            },
            "source_assembly": {
                "path": "/tmp/reviewed-signed-native-assembly.json",
                "bytes": 1,
                "sha256": "5" * 64,
            },
            "manifests": {path: {"path": path, "bytes": 1, "sha256": str(index + 4) * 64} for index, path in enumerate(manifests)},
            "vendor_tree": {"files": 3, "bytes": 3, "sha256": "8" * 64, "entries": []},
            "lock_graph": {role: {"path": value["path"], "bytes": 1, "sha256": value["candidate"]["sha256"], "graph_sha256": str(index + 9) * 64} for index, (role, value) in enumerate(role_locks.items())},
            "features": {"package": "litchi-ops", "all_features": True, "selected": ["ops"]},
        }
        identity = {
            "source_root": "",
            "endpoint": "after",
            "source_role": "candidate",
            "source_kind": "472b",
            "api_branch": "current_api",
            **fake_binding["source"],
            "manifests": fake_binding["manifests"],
            "locks": {value["path"]: {"path": value["path"], "bytes": 1, "sha256": value["candidate"]["sha256"]} for value in role_locks.values()},
            "lock_graph": fake_binding["lock_graph"],
            "vendor_tree": fake_binding["vendor_tree"],
            "toolchain": fake_binding["toolchain"],
            "probe_inputs": probes,
            "probe_provenance": {
                name: {
                    "path": str(value["path"]),
                    "bytes": value["bytes"],
                    "sha256": value["sha256"],
                }
                for name, value in probes.items()
            },
            "fixture": fixture,
            "fixture_provenance": fixture,
            "features": fake_binding["features"],
        }
        return fake_binding, identity

    def _prepared(self) -> tuple[Path, dict[str, object], mock._patch]:
        temporary = tempfile.TemporaryDirectory(prefix="signed-native-v2-consumer-")
        self.addCleanup(temporary.cleanup)
        parent = Path(temporary.name)
        source = parent / "source"
        files = {
            "Cargo.toml": b"workspace\n",
            "perf.toml": b"perf\n",
            "native.toml": b"native\n",
            "Cargo.lock": b"lock\n",
            "perf.lock": b"perf-lock\n",
            "native.lock": b"native-lock\n",
            "rust-toolchain.toml": b"toolchain\n",
            ".cargo/config.toml": b"config\n",
            "vendor/allsorts/lib": b"a",
            "vendor/flate2/lib": b"b",
            "vendor/zlib-rs/lib": b"c",
            "reviewed/probe.input": b"fn main() {}\n",
        }
        self._git_repo(source, files)
        probes, fixture = self._fixture_and_probes(source)
        binding, identity = self._fixture_patch(probes, fixture)
        identity["source_root"] = str(source.resolve())
        patches = [
            mock.patch.object(policy, "CANDIDATE_BINDING", binding),
            mock.patch.object(policy, "SOURCE_MANIFEST_PATHS", ("Cargo.toml", "perf.toml", "native.toml")),
            mock.patch.object(policy, "ROLE_LOCKS", binding["role_locks"]),
            mock.patch.object(policy, "VENDOR_ROOTS", tuple(binding["vendor_roots"])),
            mock.patch.object(policy, "TOOLCHAIN_PATHS", tuple(binding["toolchain_paths"])),
            mock.patch.object(policy, "TOOLCHAIN_IDENTITIES", binding["toolchain"]),
            mock.patch.object(policy, "PROBE_INPUTS", probes),
            mock.patch.object(policy, "FIXTURE_INPUT", fixture),
            mock.patch.object(policy, "_fixture_policy", return_value={"schema": "litchi_signed_fixture_manifest.v1", "fixtures": 1, "formats": list(policy.NATIVE_FORMATS)}),
            mock.patch.object(policy, "candidate_identity", return_value=copy.deepcopy(identity)),
        ]
        for patcher in patches:
            patcher.start()
            self.addCleanup(patcher.stop)
        output = parent / "prepared"
        result = consumer.prepare(source, output, identity=identity)
        self.assertEqual(result["status"], "prepared")
        return output, binding, mock.patch.object(policy, "candidate_identity", return_value=copy.deepcopy(identity))

    def _build_receipt(self, prepared: Path) -> tuple[Path, dict[str, object]]:
        state = consumer._revalidate_preparation(prepared)
        source_root = Path(state["source"]["source_root"])
        adapter_dependencies = sorted(consumer.ADAPTER_DIRECT_DEPENDENCIES)

        def lock_payload(include_adapter: bool) -> str:
            rows = []
            for name in adapter_dependencies:
                if name in {"serde_json", "sha2"}:
                    rows.append(
                        f'[[package]]\nname = "{name}"\nversion = "1.0.0"\n'
                        'source = "registry+https://github.com/rust-lang/crates.io-index"\n'
                        'checksum = "'
                        + ("1" if name == "serde_json" else "2")
                        + '"\n'
                    )
                else:
                    rows.append(f'[[package]]\nname = "{name}"\nversion = "0.0.0"\n')
            if include_adapter:
                rows.append(
                    '[[package]]\nname = "signed-native-consumer"\nversion = "0.0.0"\n'
                    'dependencies = [\n'
                    + "\n".join(f' "{name}",' for name in adapter_dependencies)
                    + '\n]\n'
                )
            return "\n".join(rows)

        (source_root / "Cargo.lock").write_text(lock_payload(False), encoding="utf-8")
        (prepared / "consumer/Cargo.lock").write_text(lock_payload(True), encoding="utf-8")
        binary = prepared / "target/release/signed-native-consumer"
        binary.parent.mkdir(parents=True)
        binary.write_bytes(b"not-a-native-binary")
        binary.chmod(0o755)
        build = {
            "schema": consumer.BUILD_SCHEMA,
            "status": "built",
            "preparation_receipt": state["receipt_artifact"],
            "endpoint": "after",
            "api_branch": "current_api",
            "api_feature": None,
            "source_identity": consumer._concrete_source_identity(state),
            "source_identity_receipt": state["receipt"]["source_identity"],
            "source": state["receipt"]["source"],
            "source_archive": state["receipt"]["source"]["archive"],
            "lock": state["receipt"]["source"]["locks"],
            "lock_graph": state["receipt"]["source"].get("lock_graph"),
            "source_manifests": state["receipt"]["source"].get("manifests"),
            "adapter": state["adapter"],
            "fixture": state["receipt"]["fixture"],
            "routes": state["receipt"]["routes"],
            "signed_batch": state["receipt"]["signed_batch"],
            "binary": consumer._artifact(binary, "test binary", executable=True),
            "cargo": {"path": "/tmp/cargo", "bytes": 1, "sha256": "a" * 64},
            "metadata": {"path": "/tmp/metadata", "bytes": 1, "sha256": "b" * 64},
            "consumer_lock": consumer._artifact(prepared / "consumer/Cargo.lock", "test consumer lock"),
            "consumer_lock_graph": consumer._sha256(
                consumer._canonical(list(consumer._normalized_lock_rows(prepared / "consumer/Cargo.lock", "test consumer lock")))
            ),
            "consumer_lock_binding": {
                "source_lock": state["receipt"]["source"]["locks"]["Cargo.lock"],
                "adapter_package": consumer.ADAPTER_PACKAGE,
                "adapter_external_dependencies": sorted(consumer.ADAPTER_EXTERNAL_DEPENDENCIES),
            },
        }
        build_path = prepared / "build-receipt.json"
        build_path.write_bytes(canonical(build))
        return build_path, build

    def _capabilities(self, prepared: Path, build_path: Path, build: dict[str, object]) -> Path:
        state = consumer._revalidate_preparation(prepared)
        statuses = {route: "available" for route in policy.ROUTES}
        statuses.update({route: "expected_refusal" for route in policy.CORRECTNESS_ONLY})
        build_artifact = consumer._artifact(build_path, "build receipt")
        fixture_manifest = json.loads(Path(state["fixture"]["manifest"]["path"]).read_text())
        fixture_rows = fixture_manifest["fixtures"]
        runner_artifact = consumer._artifact(consumer.RUNNER_PATH, "signed-native v2 runner")
        environment = {
            "LITCHI_SIGNED_FIXTURE_ROOT": state["fixture"]["root"],
            "LITCHI_SIGNED_FIXTURE_MANIFEST": state["fixture"]["manifest"]["path"],
            "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256": state["fixture"]["manifest"]["sha256"],
            "LITCHI_SIGNED_ENDPOINT": "after",
        }
        evidence = {}
        for index, route in enumerate(policy.ROUTES):
            results = []
            for fixture_index, fixture in enumerate(fixture_rows):
                source_digest = fixture.get("signed_sha256") or (f"{fixture_index + 1:064x}")
                result = {
                    "name": fixture["name"],
                    "format": fixture["format"],
                    "origin": fixture["origin"],
                    "signature_presence": "detected",
                    "cryptographic_verification": "not_performed",
                    "status": "expected_refusal" if route in policy.CORRECTNESS_ONLY else "ok",
                    "source_sha256": source_digest,
                }
                if route.endswith("/noop_commit"):
                    result.update({"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True, "source_after_sha256": source_digest})
                elif route.endswith("/noop_write"):
                    result.update({"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True, "source_after_sha256": source_digest, "stream_bounded": True, "writes": 1, "max_write_bytes": 1})
                elif route.endswith("/noop_replay"):
                    result.update({"patch_noop": True, "replay_base_match": True, "signed_after": True, "output_equals_source": True})
                elif route.endswith("/refusal"):
                    result.update({"refusal_kind": "ProtectedDocument", "code": "document_protected", "source_before_sha256": source_digest, "source_after_sha256": source_digest, "publication_attempts": 0, "output_published": False})
                elif route.endswith("/unsign_write"):
                    result.update({"signed_after": False, "signature_parts_absent": True, "semantic_changed": True, "stream_bounded": True, "output_equals_source": False, "signed_clone_still_refuses": True, "writes": 1, "max_write_bytes": 1})
                elif route.endswith("/unsigned_negative"):
                    result.update({"refusal_kind": "PatchConflict", "code": "invalid_input", "re_marked_signed": False, "publication_attempts": 0, "output_published": False})
                elif route.endswith("/source_freshness"):
                    result.update({"refusal_kind": "SourceChanged", "code": "io_error", "source_before_sha256": source_digest, "source_mutated_sha256": "b" * 64, "source_after_sha256": "b" * 64, "publication_attempts": 0, "output_published": False})
                elif route.endswith("/budget"):
                    result.update({"refusal_kind": "LimitExceeded", "code": "limit_exceeded", "limit_kind": "operations", "source_before_sha256": source_digest, "source_after_sha256": source_digest, "publication_attempts": 0, "output_published": False})
                results.append(result)
            result = {
                "schema": consumer.PROBE_SCHEMA,
                "route": route,
                "endpoint": "after",
                "capability": statuses[route],
                "fixtures": results,
            }
            stdout_path = prepared.parent / f"child-{index}.stdout"
            stderr_path = prepared.parent / f"child-{index}.stderr"
            stdout_path.write_bytes(canonical(result))
            stderr_path.write_bytes(b"")
            command = [str(build["binary"]["path"]), "--mode", "validate", "--route", route, "--blocks", "1"]
            child = {
                "schema": "litchi.signed-native.child.v2",
                "probe_schema": "litchi.signed-native.probe.v2",
                "mode": "validate",
                "route": route,
                "status": statuses[route],
                "endpoint": "after",
                "api_branch": build["api_branch"],
                "preparation_receipt": state["receipt_artifact"],
                "build_receipt": build_artifact,
                "source_identity": consumer._concrete_source_identity(state),
                "source_archive": state["receipt"]["source"]["archive"],
                "lock": state["receipt"]["source"]["locks"],
                "lock_graph": state["receipt"]["source"].get("lock_graph"),
                "source_manifests": state["receipt"]["source"].get("manifests"),
                "adapter": state["adapter"],
                "binary": build["binary"],
                "fixture": state["receipt"]["fixture"],
                "runner": runner_artifact,
                "argv": command,
                "command": command,
                "environment": environment,
                "cwd": str(prepared / "consumer"),
                "process": {"pid": 1000 + index, "process_group": 1000 + index, "start_new_session": True, "process_group_empty": True},
                "stdout": consumer._artifact(stdout_path, f"child {route} stdout"),
                "stderr": consumer._artifact(stderr_path, f"child {route} stderr"),
                "timeout_seconds": consumer.CHILD_TIMEOUT_SECONDS,
                "exit_code": 0,
                "timed_out": False,
                "guard_error": None,
                "after_guard_error": None,
                "result": result,
            }
            child_path = prepared.parent / f"child-{index}.json"
            child_path.write_bytes(canonical(child))
            evidence[route] = consumer._artifact(child_path, f"child {route}")
        before = prepared.parent / "batch-before.bin"
        after = prepared.parent / "batch-after.bin"
        before.write_bytes(b"signed source bytes")
        after.write_bytes(b"signed source bytes")
        batch_result = {
            "schema": consumer.BATCH_PROOF_SCHEMA,
            "selection": "immediate_refusal",
            "status": "expected_refusal",
            "error": {"kind": "DocumentProtected", "code": "document_protected", "item": None},
            "item": None,
            "iterator_reads": 0,
            "publication_attempts": 0,
            "source_before_hex": (b"signed source bytes").hex(),
            "source_after_hex": (b"signed source bytes").hex(),
            "source_bytes_equal": True,
        }
        batch_stdout = prepared.parent / "batch.stdout"
        batch_stderr = prepared.parent / "batch.stderr"
        batch_stdout.write_bytes(canonical(batch_result))
        batch_stderr.write_bytes(b"")
        value = {
            "schema": consumer.CAPABILITIES_SCHEMA,
            "protocol_schema": consumer.RUNNER_SCHEMA,
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": build_artifact,
            "endpoint": "after",
            "source_identity": consumer._concrete_source_identity(state),
            "source_identity_receipt": state["receipt"]["source_identity"],
            "source": state["receipt"]["source"],
            "source_archive": state["receipt"]["source"]["archive"],
            "lock": state["receipt"]["source"]["locks"],
            "lock_graph": state["receipt"]["source"].get("lock_graph"),
            "source_manifests": state["receipt"]["source"].get("manifests"),
            "binary": build["binary"],
            "api_branch": build["api_branch"],
            "build_receipt_binding": {
                "preparation_receipt": state["receipt_artifact"],
                "build_receipt": consumer._artifact(build_path, "build receipt"),
            },
            "adapter": state["adapter"],
            "fixture": state["receipt"]["fixture"],
            "runner": runner_artifact,
            "routes": statuses,
            "evidence": evidence,
            "signed_batch": {
                **batch_result,
                "selection": "immediate_refusal",
                "source_before": consumer._artifact(before, "batch source before"),
                "source_after": consumer._artifact(after, "batch source after"),
                "source_bytes_equal": True,
                "endpoint": "after",
                "source_identity": consumer._concrete_source_identity(state),
                "source_archive": state["receipt"]["source"]["archive"],
                "lock": state["receipt"]["source"]["locks"],
                "lock_graph": state["receipt"]["source"].get("lock_graph"),
                "source_manifests": state["receipt"]["source"].get("manifests"),
                "binary": build["binary"],
                "fixture": state["receipt"]["fixture"],
                "runner": runner_artifact,
                "argv": [str(build["binary"]["path"]), "--mode", "batch_refusal", "--blocks", "1"],
                "command": [str(build["binary"]["path"]), "--mode", "batch_refusal", "--blocks", "1"],
                "environment": environment,
                "cwd": str(prepared / "consumer"),
                "process": {"pid": 2000, "process_group": 2000, "start_new_session": True, "process_group_empty": True},
                "stdout": consumer._artifact(batch_stdout, "batch stdout"),
                "stderr": consumer._artifact(batch_stderr, "batch stderr"),
                "timeout_seconds": consumer.CHILD_TIMEOUT_SECONDS,
                "exit_code": 0,
                "timed_out": False,
                "guard_error": None,
                "after_guard_error": None,
                "preparation_receipt": state["receipt_artifact"],
                "build_receipt": build_artifact,
            },
        }
        path = prepared.parent / "capabilities.json"
        path.write_bytes(canonical(value))
        return path

    def test_positive_prepare_and_execute_consume_all_routes_and_batch_policy(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        executor_calls = []
        result = consumer.execute(
            prepared,
            executor=lambda state: (executor_calls.append(state["receipt_artifact"]) or capabilities),
        )
        self.assertEqual(result["status"], "completed")
        self.assertEqual(len(executor_calls), 1)
        execution = json.loads((prepared / "execution-receipt.json").read_text())
        self.assertEqual(tuple(execution["routes"]), policy.ROUTES)
        self.assertEqual(execution["signed_batch"]["iterator_reads"], 0)
        self.assertEqual(execution["signed_batch"]["publication_attempts"], 0)

    def test_changed_identity_and_lock_are_rejected_before_execution_output(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        recorded = json.loads((prepared / "source-identity.json").read_text())
        changed = copy.deepcopy(recorded)
        changed["commit"] = "9" * 40
        with mock.patch.object(policy, "candidate_identity", return_value=changed):
            with self.assertRaisesRegex(consumer.ConsumerError, "source identity changed"):
                consumer.execute(prepared, capabilities)
        self.assertFalse((prepared / "execution-receipt.json").exists())
        changed = copy.deepcopy(recorded)
        changed["locks"]["Cargo.lock"]["sha256"] = "a" * 64
        with mock.patch.object(policy, "candidate_identity", return_value=changed):
            with self.assertRaisesRegex(consumer.ConsumerError, "source identity changed"):
                consumer.execute(prepared, capabilities)
        self.assertFalse((prepared / "execution-receipt.json").exists())

    def test_consumer_lock_is_endpoint_bound_before_execution(self) -> None:
        prepared, _binding, _unused = self._prepared()
        self._build_receipt(prepared)
        lock_path = prepared / "consumer/Cargo.lock"
        lock_path.write_text(
            lock_path.read_text(encoding="utf-8")
            + '\n[[package]]\nname = "unreviewed-registry-package"\nversion = "9.9.9"\n',
            encoding="utf-8",
        )
        with self.assertRaisesRegex(consumer.ConsumerError, "unreviewed package"):
            consumer.execute(prepared)
        self.assertFalse((prepared / "execution-receipt.json").exists())

    def test_route_set_mismatch_and_batch_iterator_read_are_rejected(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        value["routes"][policy.ROUTES[-1]] = None
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "capability route status is invalid"):
            consumer.execute(prepared, capabilities)
        value["routes"][policy.ROUTES[-1]] = "expected_refusal"
        value["signed_batch"]["iterator_reads"] = 1
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "touched the iterator"):
            consumer.execute(prepared, capabilities)

    def test_every_route_requires_typed_evidence_and_expected_disposition(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        del value["evidence"][policy.ROUTES[0]]
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "evidence set differs"):
            consumer.execute(prepared, capabilities)
        value = json.loads(self._capabilities(prepared, build_path, build).read_text())
        value["routes"][policy.ROUTES[0]] = "capability_missing"
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "status is invalid"):
            consumer.execute(prepared, capabilities)

        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        first = policy.ROUTES[0]
        child_artifact = value["evidence"][first]
        child_path = Path(child_artifact["path"])
        child = json.loads(child_path.read_text())
        child["result"]["fixtures"][0]["source_sha256"] = "a" * 64
        stdout_path = Path(child["stdout"]["path"])
        stdout_path.write_bytes(canonical(child["result"]))
        child["stdout"] = consumer._artifact(stdout_path, "tampered fixture stdout")
        child_path.write_bytes(canonical(child))
        value["evidence"][first] = consumer._artifact(child_path, "tampered fixture source")
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "fixture source bytes differ"):
            consumer.execute(prepared, capabilities)

    def test_v1_child_and_unproven_batch_are_rejected(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        first = policy.ROUTES[0]
        child_path = Path(value["evidence"][first]["path"])
        child = json.loads(child_path.read_text())
        child["schema"] = "litchi_signed_native_child.v1"
        child_path.write_bytes(canonical(child))
        value["evidence"][first] = consumer._artifact(child_path, "tampered v1 child")
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "historical v1"):
            consumer.execute(prepared, capabilities)
        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        del value["signed_batch"]["source_before"]
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "source before"):
            consumer.execute(prepared, capabilities)

    def test_preparation_recomputes_identity_and_generated_probe(self) -> None:
        prepared, _binding, _unused = self._prepared()
        generated = (prepared / "consumer/src/main.rs").read_text()
        self.assertIn('const SCHEMA: &str = "litchi.signed-native.probe.v2";', generated)
        self.assertNotIn('const SCHEMA: &str = "litchi_signed_native_probe.v1";', generated)
        self.assertNotIn('cfg(feature = "legacy_api")', generated)
        self.assertIn("#[cfg(any())]", generated)
        self.assertIn("file.batch(&mut items, litchi_ops::BatchMode::Atomic)", generated)
        self.assertIn('pair[1] == "batch_refusal"', generated)
        source_identity = json.loads((prepared / "source-identity.json").read_text())
        source = Path(source_identity["source_root"])
        _, before_probe, _ = consumer._generated_probe(source, "before")
        self.assertNotIn('cfg(not(feature = "legacy_api"))', before_probe)
        self.assertNotIn('cfg(feature = "legacy_api")', before_probe)
        self.assertIn("edit_with_limits", before_probe)
        changed = copy.deepcopy(source_identity)
        changed["commit"] = "9" * 40
        with self.assertRaisesRegex(consumer.ConsumerError, "caller-supplied source identity"):
            consumer.prepare(source, prepared.parent / "second", identity=changed)
        # A prepared test fixture must not make the real after endpoint
        # executable while the checked-in candidate binding is pending.  The
        # three public phases all fail before creating their next artifact.
        pending_output = prepared.parent / "pending-after"
        with mock.patch.object(policy, "CANDIDATE_BINDING", None):
            with self.assertRaisesRegex(policy.IncompletePolicyError, "after endpoint is unavailable"):
                consumer.prepare(source, pending_output)
            self.assertFalse(pending_output.exists())
            with self.assertRaisesRegex(policy.IncompletePolicyError, "after endpoint is unavailable"):
                consumer.build(prepared, None)
            self.assertFalse((prepared / "target").exists())
            with self.assertRaisesRegex(policy.IncompletePolicyError, "after endpoint is unavailable"):
                consumer.execute(prepared)
            self.assertFalse((prepared / "execution-receipt.json").exists())
        (prepared / "consumer/src/main.rs").write_text("mod api_compat;\nfn main() {}\n")
        with self.assertRaisesRegex(consumer.ConsumerError, "generated signed probe"):
            consumer._revalidate_preparation(prepared)

    def test_api_branches_are_explicit_build_inputs(self) -> None:
        project = consumer._project_toml(Path("/reviewed/source"), ["docx", "pptx", "xlsx", "xlsb", "doc", "ppt", "xls"])
        self.assertIn("default = []", project)
        self.assertNotIn("legacy_api", project)
        self.assertIn('features = ["docx", "pptx", "xlsx", "xlsb", "doc", "ppt", "xls"]', project)
        self.assertEqual(consumer.API_BRANCHES, ("current_api", "legacy_api"))
        self.assertIn("commit.file;", consumer._adapter_source("before"))
        self.assertNotIn("#[cfg", consumer._adapter_source("before"))
        self.assertIn("commit.file()", consumer._adapter_source("after"))
        self.assertNotIn("legacy_api", consumer._adapter_source("after"))

    def test_runner_produces_process_bound_v2_receipts(self) -> None:
        from native_qualification import signed_native_v2_runner as runner

        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        fake = prepared.parent / "fake-signed-native"
        fake.write_text(
            """#!/usr/bin/env python3
import hashlib
import json
import os
import sys

mode = sys.argv[sys.argv.index("--mode") + 1]
if mode == "batch_refusal":
    payload = b"signed source bytes"
    print(json.dumps({
        "schema": "litchi.signed-native.batch-proof.v2",
        "selection": "immediate_refusal",
        "status": "expected_refusal",
        "error": {"kind": "DocumentProtected", "code": "document_protected", "item": None},
        "item": None,
        "iterator_reads": 0,
        "publication_attempts": 0,
        "source_before_hex": payload.hex(),
        "source_after_hex": payload.hex(),
        "source_bytes_equal": True,
    }))
    raise SystemExit(0)
route = sys.argv[sys.argv.index("--route") + 1]
refusal = any(route.endswith("/" + suffix) for suffix in ("refusal", "unsigned_negative", "source_freshness", "budget"))
status = "expected_refusal" if refusal else "available"
rows = []
with open(os.environ["LITCHI_SIGNED_FIXTURE_MANIFEST"], encoding="utf-8") as manifest:
    fixtures = json.load(manifest)["fixtures"]
for index, fixture in enumerate(fixtures):
    source = fixture.get("signed_sha256") or (f"{index + 1:064x}")
    row = {
        "name": fixture["name"], "format": fixture["format"], "origin": fixture["origin"],
        "signature_presence": "detected", "cryptographic_verification": "not_performed",
        "status": "expected_refusal" if refusal else "ok", "source_sha256": source,
    }
    if route.endswith("/noop_commit"):
        row.update({"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True, "source_after_sha256": source})
    elif route.endswith("/noop_write"):
        row.update({"diagnostics_changed": False, "patch_noop": True, "signed_after": True, "output_equals_source": True, "source_after_sha256": source, "stream_bounded": True, "writes": 1, "max_write_bytes": 1})
    elif route.endswith("/noop_replay"):
        row.update({"patch_noop": True, "replay_base_match": True, "signed_after": True, "output_equals_source": True})
    elif route.endswith("/refusal"):
        row.update({"refusal_kind": "ProtectedDocument", "code": "document_protected", "source_before_sha256": source, "source_after_sha256": source, "publication_attempts": 0, "output_published": False})
    elif route.endswith("/unsign_write"):
        row.update({"signed_after": False, "signature_parts_absent": True, "semantic_changed": True, "stream_bounded": True, "output_equals_source": False, "signed_clone_still_refuses": True, "writes": 1, "max_write_bytes": 1})
    elif route.endswith("/unsigned_negative"):
        row.update({"refusal_kind": "PatchConflict", "code": "invalid_input", "re_marked_signed": False, "publication_attempts": 0, "output_published": False})
    elif route.endswith("/source_freshness"):
        row.update({"refusal_kind": "SourceChanged", "code": "io_error", "source_before_sha256": source, "source_mutated_sha256": "b" * 64, "source_after_sha256": "b" * 64, "publication_attempts": 0, "output_published": False})
    elif route.endswith("/budget"):
        row.update({"refusal_kind": "LimitExceeded", "code": "limit_exceeded", "limit_kind": "operations", "source_before_sha256": source, "source_after_sha256": source, "publication_attempts": 0, "output_published": False})
    rows.append(row)
print(json.dumps({"schema": "litchi.signed-native.probe.v2", "route": route, "capability": status, "fixtures": rows}))
""",
            encoding="utf-8",
        )
        fake.chmod(0o755)
        build["binary"] = consumer._artifact(fake, "fake signed-native binary", executable=True)
        build_path.write_bytes(canonical(build))
        capabilities = runner.run_prepared(prepared, output=prepared.parent / "runner-output")
        result = consumer.execute(prepared, capabilities)
        self.assertEqual(result["status"], "completed")
        value = json.loads(capabilities.read_text())
        self.assertEqual(value["protocol_schema"], consumer.RUNNER_SCHEMA)
        self.assertEqual(len(value["evidence"]), len(policy.ROUTES))
        self.assertEqual(value["signed_batch"]["iterator_reads"], 0)


if __name__ == "__main__":
    unittest.main()
