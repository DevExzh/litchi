"""Regression tests for ordinary and original-to-final W8 analysis records."""

from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest import mock

from tools import w8_file_provenance as provenance


ROOT = Path(__file__).resolve().parents[1]
ANALYZER_PATH = ROOT / "docs/unified-ops/results/0362h/analyze.py"
PROBE_PATH = ROOT / "crates/litchi-ops/examples/file_w8_qualified_probe.rs"


def load_analyzer():
    spec = importlib.util.spec_from_file_location("w8_pair_analyzer", ANALYZER_PATH)
    if spec is None or spec.loader is None:
        raise AssertionError(f"cannot load {ANALYZER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class PairAnalyzerTests(unittest.TestCase):
    def _fixture(self, root: Path) -> tuple[dict, object, dict[str, Path]]:
        analyzer = load_analyzer()
        probe = root / "probe.rs"
        probe.write_bytes(PROBE_PATH.read_bytes())
        fixture = root / "fixture.json"
        fixture.write_bytes(provenance.canonical(provenance.fixture_descriptor(probe)))
        fixture_row = provenance.artifact(fixture)
        probe_row = provenance.artifact(probe)
        routes = provenance.fixture_routes(json.loads(fixture.read_bytes()))
        host = root / "host.json"
        host.write_bytes(provenance.canonical({"blockers": [], "processes": [], "allowed": []}))
        stdout = root / "stdout"
        stderr = root / "stderr"
        stdout.write_bytes(b"probe stdout\n")
        stderr.write_bytes(b"")
        host_row = provenance.artifact(host)
        stdout_row = provenance.artifact(stdout)
        stderr_row = provenance.artifact(stderr)

        binary_rows = []
        endpoint_values = []
        fixed_inputs = {}
        for side, commit in (("before", "a" * 40), ("after", "b" * 40)):
            side_root = root / side
            side_root.mkdir()
            binary = side_root / "probe"
            binary.write_bytes(f"{side} binary\n".encode())
            os.chmod(binary, 0o755)
            source = side_root / "source.tar"
            lock = side_root / "Cargo.lock"
            workspace = side_root / "Cargo.toml"
            package = side_root / "litchi-ops.Cargo.toml"
            toolchain = side_root / "rust-toolchain.toml"
            source.write_bytes(f"{side} source\n".encode())
            lock.write_bytes(f"{side} lock\n".encode())
            workspace.write_bytes(f"{side} workspace\n".encode())
            package.write_bytes(f"{side} package\n".encode())
            toolchain.write_bytes(f"{side} toolchain\n".encode())
            projection = side_root / "projection.json"
            projection.write_bytes(
                provenance.canonical(
                    {"schema": "litchi_w8_dependency_projection.v1", "side": side}
                )
            )
            source_assembly = side_root / "source-assembly.json"
            build = side_root / "build.json"
            validation = side_root / "validation.json"
            source_assembly.write_bytes(provenance.canonical({"side": side, "kind": "none"}))
            build.write_bytes(provenance.canonical({"side": side, "status": "complete"}))
            validation.write_bytes(provenance.canonical({"side": side, "status": "complete"}))
            artifacts = {
                "binary": provenance.artifact(binary),
                "source": provenance.artifact(source),
                "lock": provenance.artifact(lock),
                "probe": probe_row,
                "fixture": fixture_row,
                "dependency_projection": provenance.artifact(projection),
                "workspace_manifest": provenance.artifact(workspace),
                "package_manifest": provenance.artifact(package),
                "toolchain": provenance.artifact(toolchain),
                "source_assembly": provenance.artifact(source_assembly),
                "build_receipt": provenance.artifact(build),
                "validation": provenance.artifact(validation),
            }
            endpoint_path = side_root / "endpoint.json"
            endpoint_value = {
                "schema": "litchi_file_endpoint.v1",
                "source_commit": commit,
                "dependency_projection_sha256": artifacts["dependency_projection"]["sha256"],
                **artifacts,
                "features": {"package": "litchi-ops", "selected": ["all"], "all_features": True},
                "target": "test-target",
                "profile": "release",
            }
            endpoint_path.write_bytes(provenance.canonical(endpoint_value))
            endpoint_value["manifest"] = provenance.artifact(endpoint_path)
            endpoint_value["routes"] = routes
            endpoint_values.append(endpoint_value)
            binary_rows.append(artifacts["binary"])
            fixed_inputs[side] = {
                "source_commit": commit,
                "source": artifacts["source"],
                "workspace_manifest": artifacts["workspace_manifest"],
                "package_manifest": artifacts["package_manifest"],
                "toolchain": artifacts["toolchain"],
                "lock": artifacts["lock"],
                "projection": artifacts["dependency_projection"],
            }

        proof = {
            "registry_package_count": 0,
            "registry_identities_checksums_dependencies_equal": True,
            "local_edge_differences": [],
        }
        policy_identity = {
            "schema": "test.original_to_final_policy.v1",
            "version": "test-v1",
            "sha256": "c" * 64,
        }
        fake_policy = types.SimpleNamespace(
            PAIR_SCHEMA="test.original_to_final_pair.v1",
            REVIEWED_POLICY_SHA256=policy_identity["sha256"],
            FIXED_INPUTS=fixed_inputs,
            PROBE_BYTES=probe_row["bytes"],
            PROBE_SHA256=probe_row["sha256"],
            FIXTURE_BYTES=fixture_row["bytes"],
            FIXTURE_SHA256=fixture_row["sha256"],
            TARGET="test-target",
            PROFILE="release",
            FEATURES=("all",),
            TOOLCHAIN_SHA256=fixed_inputs["before"]["toolchain"]["sha256"],
            load_reviewed_policy=lambda: {
                "schema": policy_identity["schema"],
                "policy_version": policy_identity["version"],
            },
            validate_projection_pair=lambda _before, _after: proof,
        )
        before_endpoint, after_endpoint = endpoint_values
        workload = {
            "probe_sha256": probe_row["sha256"],
            "fixture_sha256": fixture_row["sha256"],
            "route_count": len(routes),
            "routes": routes,
        }
        build = {
            "target": "test-target",
            "profile": "release",
            "features": ["all"],
            "toolchain_sha256": fake_policy.TOOLCHAIN_SHA256,
        }
        receipt_value = {
            "schema": fake_policy.PAIR_SCHEMA,
            "policy": policy_identity,
            "before": analyzer._pair_endpoint_part(before_endpoint, "before"),
            "after": analyzer._pair_endpoint_part(after_endpoint, "after"),
            "workload": workload,
            "build": build,
            "projection": proof,
            "status": "verified_protocol_only",
        }
        receipt_path = root / "pair-receipt.json"
        receipt_path.write_bytes(provenance.canonical(receipt_value))
        shared = {
            "mode": "original_to_final",
            "policy": policy_identity,
            "receipt": provenance.artifact(receipt_path),
            "projection": proof,
            "workload": workload,
            "build": build,
        }
        return shared, fake_policy, {
            "before": before_endpoint,
            "after": after_endpoint,
            "routes": routes,
            "binary_before": Path(binary_rows[0]["path"]),
            "binary_after": Path(binary_rows[1]["path"]),
            "source_before": Path(before_endpoint["source"]["path"]),
            "projection_before": Path(before_endpoint["dependency_projection"]["path"]),
            "lock_before": Path(before_endpoint["lock"]["path"]),
            "host": host_row,
            "stdout": stdout_row,
            "stderr": stderr_row,
        }

    @staticmethod
    def _phase() -> dict:
        return {"samples": [150_000_000], "calls": 1, "total_ns": 150_000_000, "operation_count": 1}

    def _write_analysis_manifest(
        self,
        root: Path,
        shared: dict,
        endpoint_values: dict,
    ) -> Path:
        changed_suffixes = {
            "set_commit", "set_publication", "batch_commit", "batch_publication",
            "import_commit", "import_publication", "merge_commit", "merge_publication",
        }
        routes = endpoint_values["routes"]
        before_binary = provenance.artifact(endpoint_values["binary_before"])
        after_binary = provenance.artifact(endpoint_values["binary_after"])
        receipt_root = root / "receipts"
        receipt_root.mkdir()
        unique = 0

        def probe_receipt(side: str, route: str, mode: str, calls: int | None) -> dict:
            nonlocal unique
            unique += 1
            binary = before_binary if side == "before" else after_binary
            label = f"synthetic-{unique}"
            path = receipt_root / f"{label}.receipt.json"
            command = [binary["path"], "--mode", mode, "--route", route]
            receipt = {
                "schema": "litchi_file_probe_receipt.v2",
                "label": label,
                "binary": binary,
                "route": route,
                "mode": mode,
                "calls": calls,
                "command": command,
                "stdout": endpoint_values["stdout"],
                "stderr": endpoint_values["stderr"],
                "exit_code": 0,
                "overlap": None,
                "guard_error": None,
                "host_snapshots": [endpoint_values["host"]],
            }
            path.write_bytes(provenance.canonical(receipt))
            return {"_receipt": str(path), "_host_snapshots": [endpoint_values["host"]]}

        def semantic_value(route: str) -> dict:
            suffix = route.rsplit("/", 1)[-1]
            changed = suffix in changed_suffixes
            output = {
                "bytes": 1,
                "sha256": "2" * 64,
                "source_sha256": "1" * 64,
            }
            validation = {
                "semantic_state_valid": True,
                "source_state_preserved": True,
                "semantic_equal": True,
                "source_state_exact": True,
                "leaf_repeated_equal": True,
                "facade_repeated_equal": True,
                "exact_noop": True,
                "leaf_changed": changed,
                "facade_changed": changed,
                "leaf_output_changed": changed,
                "facade_output_changed": changed,
                "source_leaf_output": output,
                "source_facade_output": output,
                "leaf_output": output,
                "facade_output": output,
                "leaf_repeat": output,
                "facade_repeat": output,
                "source_leaf_before_sha256": "3" * 64,
                "source_leaf_after_sha256": "3" * 64,
                "source_facade_before_sha256": "4" * 64,
                "source_facade_after_sha256": "4" * 64,
            }
            if suffix == "open_get":
                validation.update({
                    "read_marker_equal": True,
                    "leaf_read_marker_len": 1,
                    "facade_read_marker_len": 1,
                })
            return validation

        route_lists = {
            "before": probe_receipt("before", "*", "list", None),
            "after": probe_receipt("after", "*", "list", None),
        }
        calibration = []
        for route in routes:
            rows = []
            for side in ("before", "after"):
                rows.append({
                    "common_calls": 1,
                    "minimum_elapsed_ns": 150_000_000,
                    **probe_receipt(side, route, "calibrate", None),
                })
            calibration.append({"route": route, "common_calls": 1, "rows": rows})

        abba = []
        for route in routes:
            blocks = []
            for block_index in range(5):
                legs = []
                for side in ("before", "after", "after", "before"):
                    value = {
                        "route": route,
                        "calls": 1,
                        **probe_receipt(side, route, "measure", 1),
                        "validation": semantic_value(route),
                        "blocks": [{
                            "order": ["leaf", "facade", "facade", "leaf"],
                            "leaf": [self._phase(), self._phase()],
                            "facade": [self._phase(), self._phase()],
                        }],
                    }
                    legs.append({"side": side, "value": value})
                blocks.append({"block": block_index, "legs": legs})
            abba.append({"route": route, "calls": 1, "blocks": blocks})

        manifest = {
            "schema": "litchi_file_qualification.v3",
            "before": before_binary,
            "after": after_binary,
            "routes": routes,
            "route_lists": route_lists,
            "rounds": 5,
            "target_ns": 500_000_000,
            "minimum_sample_ns": 150_000_000,
            "provenance": {
                "before": endpoint_values["before"],
                "after": endpoint_values["after"],
                "shared": shared,
            },
            "calibration": calibration,
            "abba": abba,
        }
        output = root / "analysis-output"
        output.mkdir()
        (output / "manifest.json").write_bytes(provenance.canonical(manifest))
        return output

    def test_pair_analysis_accepts_exact_record_and_rejects_source_projection_lock_tampering(self) -> None:
        analyzer = load_analyzer()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shared, fake_policy, values = self._fixture(root)
            output = self._write_analysis_manifest(root, shared, values)
            policy_patch = mock.patch.object(analyzer, "_pair_policy_module", return_value=fake_policy)
            with policy_patch, mock.patch.object(
                sys, "argv", [str(ANALYZER_PATH), str(output)]
            ):
                analyzer.main()
            self.assertTrue((output / "analysis.json").is_file())

            for label, path, payload in (
                ("source", values["source_before"], b"tampered source\n"),
                ("projection", values["projection_before"], b"tampered projection\n"),
                ("lock", values["lock_before"], b"tampered lock\n"),
            ):
                with self.subTest(label=label):
                    original = path.read_bytes()
                    path.write_bytes(payload)
                    try:
                        with policy_patch, mock.patch.object(
                            sys, "argv", [str(ANALYZER_PATH), str(output)]
                        ), self.assertRaises(SystemExit):
                            analyzer.main()
                    finally:
                        path.write_bytes(original)


if __name__ == "__main__":
    unittest.main()
