"""Source-only contract tests for the supplemental cold PPT route."""
from __future__ import annotations
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
from typing import Mapping, Sequence
import unittest
from unittest import mock
ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "ppt_lazy_cold_first_commit", ROOT / "tools/ppt_lazy_cold_first_commit.py"
)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
class ColdFirstCommitTests(unittest.TestCase):
    def test_route_is_separate_and_cold(self) -> None:
        self.assertEqual(MODULE.ROUTE, "primary_first_commit_cold")
        self.assertEqual(MODULE.UNDERLYING_ROUTE, "primary_noop_commit")
        self.assertEqual(MODULE.CPU, 29)
        self.assertIn("--warmups", MODULE.ARGS)
        self.assertEqual(MODULE.ARGS[MODULE.ARGS.index("--warmups") + 1], "0")
        self.assertEqual(MODULE.ARGS[MODULE.ARGS.index("--iterations") + 1], "1")
    def test_source_only_plan_binds_reviewed_binary_and_fixture(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-cold-plan-") as temporary:
            root = Path(temporary)
            binary = root / "ppt_lazy_native"
            fixture = root / "fixture.ppt"
            binary.write_bytes(b"reviewed binary")
            fixture.write_bytes(b"fixture")
            # Use the fixed identity table only to test that callers cannot
            # supply an arbitrary expected hash; no execution is performed.
            role = "candidate-normal"
            expected_binary = MODULE.BINARY_IDENTITIES[role]
            binary.write_bytes(b"x" * expected_binary["bytes"])
            os.chmod(binary, 0o755)
            digest = hashlib.sha256(binary.read_bytes()).hexdigest()
            MODULE.BINARY_IDENTITIES[role]["sha256"] = digest
            MODULE.FIXTURE["bytes"] = fixture.stat().st_size
            MODULE.FIXTURE["sha256"] = hashlib.sha256(fixture.read_bytes()).hexdigest()
            try:
                result = MODULE.run(binary, fixture, role, execute=False)
                self.assertEqual(result["route"], "primary_first_commit_cold")
                self.assertEqual(result["command"][result["command"].index("--warmups") + 1], "0")
                self.assertNotIn("report", result)
            finally:
                MODULE.BINARY_IDENTITIES[role]["sha256"] = "ac72404c98bf37f5a7b9433c01880547a4b34d24a8d346dc359ecc765770ff1a"
                MODULE.FIXTURE["bytes"] = 14_848
                MODULE.FIXTURE["sha256"] = "02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c"
    def test_report_requires_one_cold_iteration_and_exact_noop(self) -> None:
        fixture = dict(MODULE.FIXTURE)
        proof = {
            "output": fixture,
            "source_exact": True,
            "changed": False,
            "patch_noop": True,
            "operation_count": 0,
            "patch_before_exact": True,
            "patch_after_exact": True,
            "entry": "commit",
            "fixture": {"key": "default", "output": fixture},
        }
        report = {
            "schema": MODULE.PROBE_SCHEMA,
            "mode": "timed",
            "route": MODULE.UNDERLYING_ROUTE,
            "samples": [{
                "elapsed_ns": 1,
                "iterations": 1,
                "allocation": {"status": "unavailable", "scope": "operation_global_system_allocator"},
            }],
            "fixture": {"key": "default", "output": fixture},
            "postcondition": proof,
        }
        self.assertEqual(MODULE.validate_report(report, fixture, allocator=False)["route"], MODULE.UNDERLYING_ROUTE)
        report["samples"][0]["iterations"] = 2
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_report(report, fixture, allocator=False)
    def test_strict_json_rejects_duplicate_and_nonfinite_values(self) -> None:
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE._strict_json(b'{"x": 1, "x": 2}', "test")
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE._strict_json(b'{"x": 1e999}', "test")

    def test_public_cli_exposes_symmetric_pair_contract(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(ROOT / "tools" / "ppt_lazy_cold_first_commit.py"), "--help"],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(completed.returncode, 0)
        self.assertIn("--manifest", completed.stdout)
        self.assertIn("--projection", completed.stdout)
        self.assertIn("--out-dir", completed.stdout)
        self.assertIn("--rounds", completed.stdout)

    def test_pair_report_aggregates_operation_elapsed_ns_in_abba_order(self) -> None:
        rows = []
        for round_index in range(1):
            for role in MODULE.ABBA_ORDER:
                for kind in ("normal", "allocator"):
                    rows.append({
                        "round": round_index,
                        "role": role,
                        "endpoint": MODULE.ROLE_ENDPOINT[role],
                        "kind": kind,
                        "sample": {
                            "elapsed_ns": 10 if kind == "normal" else 20,
                            "iterations": 1,
                            "allocation": {"status": "unavailable"},
                        },
                    })
        report = {
            "schema": MODULE.SCHEMA,
            "route": MODULE.ROUTE,
            "abba_order": list(MODULE.ABBA_ORDER),
            "rounds": 1,
            "samples": rows,
            "aggregate": {
                "baseline": {
                    "normal": {"operation_total_ns": 20, "operation_count": 2, "operation_mean_ns": 10},
                    "allocator": {"operation_total_ns": 40, "operation_count": 2, "operation_mean_ns": 20},
                },
                "candidate": {
                    "normal": {"operation_total_ns": 20, "operation_count": 2, "operation_mean_ns": 10},
                    "allocator": {"operation_total_ns": 40, "operation_count": 2, "operation_mean_ns": 20},
                },
            },
        }
        for row in rows:
            row["event"] = {
                "label": f"cold-{row['kind']}-round-{row['round']}-{row['role']}",
                "endpoint": row["endpoint"],
                "binary": {"path": "/tmp/reviewed-binary", "bytes": 1, "sha256": "0" * 64},
                "args": list(MODULE.ARGS),
                "guard": {
                    "before": {},
                    "live": [],
                    "after": {},
                    "memory_limit_bytes": 1,
                    "max_output_bytes": 1,
                    "max_load": 1.0,
                },
                "target": {"path": "/tmp/target", "bytes_before": 0, "bytes_after": 0, "hard_limit_bytes": 1},
                "event": {
                    "label": f"cold-{row['kind']}-round-{row['round']}-{row['role']}",
                    "command": [],
                    "cwd": "/tmp",
                    "returncode": 0,
                    "wall_ns": 1,
                    "target": {},
                    "stdout": {},
                    "stderr": {},
                    "monitor": {},
                },
            }
        self.assertNotIn("median", MODULE.validate_pair_report(report, rounds=1))
        full = dict(
            report,
            underlying_route=MODULE.UNDERLYING_ROUTE,
            warmups=0,
            iterations=1,
            commands={},
            manifest={"path": "manifest"},
            build_input_projection={"path": "projection"},
            claim_scope="operation only",
            execution="native_executed",
        )
        self.assertEqual(MODULE.validate_pair_report(full, rounds=1)["execution"], "native_executed")
        reordered = dict(report, samples=list(report["samples"]))
        reordered["samples"][0], reordered["samples"][2] = reordered["samples"][2], reordered["samples"][0]
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_pair_report(reordered, rounds=1)
        report["aggregate"]["candidate"]["normal"]["operation_total_ns"] = 1
        with self.assertRaises(MODULE.ColdRouteError):
            MODULE.validate_pair_report(report, rounds=1)

    def test_execute_pair_appends_and_validates_public_event_provenance(self) -> None:
        fixture = dict(MODULE.FIXTURE)
        proof = {
            "output": fixture,
            "source_exact": True,
            "changed": False,
            "patch_noop": True,
            "operation_count": 0,
            "patch_before_exact": True,
            "patch_after_exact": True,
            "entry": "commit",
            "fixture": {"key": "default", "output": fixture},
        }

        class FakePublicRunner:
            def __init__(self, root: Path) -> None:
                self.root = root
                self.rechecks: list[str] = []
                self.event_checks = 0
                self.child_labels: list[str] = []
                self.endpoints = {
                    name: self._endpoint(name)
                    for name in ("baseline", "candidate")
                }

            def _endpoint(self, name: str) -> dict[str, object]:
                root = self.root / f"{name}-source"
                target = self.root / f"{name}-normal-target"
                allocator_target = self.root / f"{name}-allocator-target"
                root.mkdir()
                target.mkdir()
                allocator_target.mkdir()
                return {
                    "name": name,
                    "root": root,
                    "target_dir": target,
                    "allocator_target_dir": allocator_target,
                    "normal_binary": {
                        "path": str(target / "ppt_lazy_native"),
                        "bytes": 1,
                        "sha256": "0" * 64,
                    },
                    "allocator_binary": {
                        "path": str(allocator_target / "ppt_lazy_native_alloc"),
                        "bytes": 1,
                        "sha256": "0" * 64,
                    },
                }

            def _load_manifest(self, path: Path, *, plan: bool) -> dict[str, object]:
                return {
                    "build": {"fixtures": {"default": fixture}},
                    "protocol": {"cpu": MODULE.CPU, "max_load": 10.0, "timeout_s": 1.0},
                }

            def validate_inputs(self, manifest: Path, projection: Path) -> dict[str, object]:
                return {
                    "manifest": {"path": str(manifest), "bytes": 1, "sha256": "1" * 64},
                    "build_input_projection": {"path": str(projection), "bytes": 1, "sha256": "2" * 64},
                    "baseline": self.endpoints["baseline"],
                    "candidate": self.endpoints["candidate"],
                    "probe_sources": [],
                    "build_receipt": {},
                }

            @staticmethod
            def _overlap(first: Path, second: Path) -> bool:
                try:
                    first.relative_to(second)
                    return True
                except ValueError:
                    try:
                        second.relative_to(first)
                        return True
                    except ValueError:
                        return False

            @staticmethod
            def _prepare_output(path: Path, roots: Sequence[Path]) -> Path:
                path.mkdir(parents=True)
                (path / "tmp").mkdir()
                return path

            def _recheck(self, endpoint: Mapping[str, object], *args: object) -> dict[str, object]:
                self.rechecks.append(str(endpoint["name"]))
                return {"phase": str(args[3])}

            def _run_child(
                self,
                endpoint: Mapping[str, object],
                protocol: Mapping[str, object],
                out_dir: Path,
                label: str,
                binary: Path,
                target_dir: Path,
                args: Sequence[str],
            ) -> tuple[dict[str, object], dict[str, object]]:
                self.child_labels.append(label)
                allocator = binary.name.endswith("_alloc")
                allocation = (
                    {
                        "status": "measured",
                        "scope": "operation_global_system_allocator",
                        "allocation_calls": 0,
                        "deallocation_calls": 0,
                        "reallocation_calls": 0,
                        "failed_allocation_calls": 0,
                        "allocated_bytes": 0,
                        "deallocated_bytes": 0,
                        "live_bytes_before": 0,
                        "live_bytes_after": 0,
                        "peak_live_bytes_before": 0,
                        "peak_live_bytes_after": 0,
                    }
                    if allocator
                    else {"status": "unavailable", "scope": "operation_global_system_allocator"}
                )
                report = {
                    "schema": MODULE.PROBE_SCHEMA,
                    "mode": "timed",
                    "route": MODULE.UNDERLYING_ROUTE,
                    "fixture": {"key": "default", "output": fixture},
                    "postcondition": proof,
                    "samples": [{"elapsed_ns": 11, "iterations": 1, "allocation": allocation}],
                }
                child_event = {
                    "label": label,
                    "command": [str(binary), *args],
                    "cwd": str(endpoint["root"]),
                    "returncode": 0,
                    "wall_ns": 11,
                    "target": {},
                    "stdout": {},
                    "stderr": {},
                    "monitor": {},
                }
                event = {
                    "label": label,
                    "endpoint": endpoint["name"],
                    "binary": {"path": str(binary), "bytes": 1, "sha256": "0" * 64},
                    "args": list(args),
                    "guard": {
                        "before": {},
                        "live": [],
                        "after": {},
                        "memory_limit_bytes": 1,
                        "max_output_bytes": 1,
                        "max_load": 10.0,
                    },
                    "target": {"path": str(target_dir), "bytes_before": 0, "bytes_after": 0, "hard_limit_bytes": 1},
                    "event": child_event,
                }
                return report, event

            def _validate_child_event(self, event: Mapping[str, object], **kwargs: object) -> None:
                self.event_checks += 1
                if event["label"] != kwargs["label"]:
                    raise AssertionError("fake public event label mismatch")

        with tempfile.TemporaryDirectory(prefix="ppt-cold-public-execute-") as temporary:
            root = Path(temporary)
            manifest = root / "manifest.json"
            projection = root / "projection.json"
            manifest.write_text("manifest", encoding="utf-8")
            projection.write_text("projection", encoding="utf-8")
            fake = FakePublicRunner(root)
            output = root / "output"
            with mock.patch.object(MODULE, "_public_runner", return_value=fake):
                result = MODULE.run_pair(manifest, projection, output, rounds=1, execute=True)
            self.assertEqual(result["execution"], "native_executed")
            self.assertEqual(len(result["samples"]), 8)
            self.assertEqual(fake.event_checks, 8)
            self.assertEqual(len(fake.rechecks), 8)
            self.assertEqual(MODULE.validate_pair_report(result, rounds=1)["aggregate"]["candidate"]["normal"]["operation_count"], 2)
            tampered = json.loads(json.dumps(result))
            del tampered["samples"][0]["event"]
            with self.assertRaises(MODULE.ColdRouteError):
                MODULE.validate_pair_report(tampered, rounds=1)
if __name__ == "__main__":
    unittest.main()
