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
                    "relative_path": fixture_relative,
                    "input_sha256": hashlib.sha256(fixture_payload).hexdigest(),
                }
            ],
        }
        fixture_path = source / "fixtures.json"
        fixture_path.write_bytes(canonical(fixture_value))
        probes: dict[str, dict[str, object]] = {}
        for index, name in enumerate(("probe", "qualification_driver", "validator", "memory_collector", "final_collector")):
            relative = f"reviewed/{name}.input"
            payload = ("fn main() {}\n" if name == "probe" else f"input-{index}\n").encode()
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
            "manifests": {path: {"path": path, "bytes": 1, "sha256": str(index + 4) * 64} for index, path in enumerate(manifests)},
            "vendor_tree": {"files": 3, "bytes": 3, "sha256": "8" * 64, "entries": []},
            "lock_graph": {role: {"path": value["path"], "bytes": 1, "sha256": value["candidate"]["sha256"], "graph_sha256": str(index + 9) * 64} for index, (role, value) in enumerate(role_locks.items())},
            "features": {"package": "litchi-ops", "all_features": True, "selected": ["ops"]},
        }
        identity = {
            "source_root": "",
            **fake_binding["source"],
            "manifests": fake_binding["manifests"],
            "locks": {value["path"]: {"path": value["path"], "bytes": 1, "sha256": value["candidate"]["sha256"]} for value in role_locks.values()},
            "lock_graph": fake_binding["lock_graph"],
            "vendor_tree": fake_binding["vendor_tree"],
            "toolchain": fake_binding["toolchain"],
            "probe_inputs": probes,
            "fixture": fixture,
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
        binary = prepared / "target/release/signed-native-consumer"
        binary.parent.mkdir(parents=True)
        binary.write_bytes(b"not-a-native-binary")
        binary.chmod(0o755)
        build = {
            "schema": consumer.BUILD_SCHEMA,
            "status": "built",
            "preparation_receipt": state["receipt_artifact"],
            "source_identity": consumer._concrete_source_identity(state),
            "source_identity_receipt": state["receipt"]["source_identity"],
            "source": state["receipt"]["source"],
            "source_archive": state["receipt"]["source"]["archive"],
            "lock": state["receipt"]["source"]["locks"],
            "adapter": state["adapter"],
            "fixture": state["receipt"]["fixture"],
            "routes": state["receipt"]["routes"],
            "signed_batch": state["receipt"]["signed_batch"],
            "binary": consumer._artifact(binary, "test binary", executable=True),
            "cargo": {"path": "/tmp/cargo", "bytes": 1, "sha256": "a" * 64},
            "metadata": {"path": "/tmp/metadata", "bytes": 1, "sha256": "b" * 64},
        }
        build_path = prepared / "build-receipt.json"
        build_path.write_bytes(canonical(build))
        return build_path, build

    def _capabilities(self, prepared: Path, build_path: Path, build: dict[str, object]) -> Path:
        state = consumer._revalidate_preparation(prepared)
        statuses = {route: "available" for route in policy.ROUTES}
        value = {
            "schema": consumer.CAPABILITIES_SCHEMA,
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": consumer._artifact(build_path, "build receipt"),
            "source_identity": consumer._concrete_source_identity(state),
            "source_identity_receipt": state["receipt"]["source_identity"],
            "source": state["receipt"]["source"],
            "source_archive": state["receipt"]["source"]["archive"],
            "lock": state["receipt"]["source"]["locks"],
            "binary": build["binary"],
            "build_receipt_binding": {
                "preparation_receipt": state["receipt_artifact"],
                "build_receipt": consumer._artifact(build_path, "build receipt"),
            },
            "adapter": state["adapter"],
            "fixture": state["receipt"]["fixture"],
            "routes": statuses,
            "signed_batch": {
                "selection": "immediate_refusal",
                "status": "expected_refusal",
                "iterator_reads": 0,
                "publication_attempts": 0,
            },
        }
        path = prepared.parent / "capabilities.json"
        path.write_bytes(canonical(value))
        return path

    def test_positive_prepare_and_execute_consume_all_routes_and_batch_policy(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        result = consumer.execute(prepared, capabilities)
        self.assertEqual(result["status"], "completed")
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

    def test_route_set_mismatch_and_batch_iterator_read_are_rejected(self) -> None:
        prepared, _binding, _unused = self._prepared()
        build_path, build = self._build_receipt(prepared)
        capabilities = self._capabilities(prepared, build_path, build)
        value = json.loads(capabilities.read_text())
        value["routes"][policy.ROUTES[-1]] = None
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "capability route status is invalid"):
            consumer.execute(prepared, capabilities)
        value["routes"][policy.ROUTES[-1]] = "available"
        value["signed_batch"]["iterator_reads"] = 1
        capabilities.write_bytes(canonical(value))
        with self.assertRaisesRegex(consumer.ConsumerError, "touched the iterator"):
            consumer.execute(prepared, capabilities)


if __name__ == "__main__":
    unittest.main()
