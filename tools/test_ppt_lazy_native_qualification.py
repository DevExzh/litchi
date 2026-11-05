import json
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

try:
    from tools import ppt_lazy_native_qualification as protocol
except ImportError:  # Support direct execution from the tools directory.
    import ppt_lazy_native_qualification as protocol


ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "tools" / "native_qualification" / "manifests" / "ppt-lazy-edit.json"


class PptLazyNativeProtocolTests(unittest.TestCase):
    def test_public_cli_binds_canonical_build_input_projection(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-public-plan-") as temporary:
            root = Path(temporary)
            manifest_path = root / "manifest.json"
            projection_path = root / "projection.json"
            output = root / "output"
            manifest_path.write_bytes(protocol._canonical_json(manifest))
            projection_path.write_bytes(
                protocol._canonical_json(protocol._build_input_projection_value(manifest))
            )
            self.assertEqual(
                protocol.main(
                    [
                        "--manifest", str(manifest_path),
                        "--build-input-projection", str(projection_path),
                        "--plan", "--out-dir", str(output),
                    ]
                ),
                0,
            )
            self.assertTrue((output / "plan.json").is_file())
            changed = dict(manifest)
            changed["protocol"] = dict(manifest["protocol"], cpu=29)
            projection_path.write_bytes(
                protocol._canonical_json(protocol._build_input_projection_value(changed))
            )
            with self.assertRaises(protocol.RunFailure):
                protocol._validate_build_input_projection(manifest, manifest_path, projection_path)

    def test_public_endpoint_lifecycle_rejects_target_lock_binary_and_worktree_mutations(self):
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-endpoint-") as temporary:
            outer = Path(temporary)
            root = outer / "candidate"
            root.mkdir()
            subprocess.run(["git", "-C", str(root), "init", "--quiet"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.name", "PPT test"], check=True)
            lock = root / "tools" / "perf-baseline" / "Cargo.lock"
            lock.parent.mkdir(parents=True)
            lock.write_bytes(b"candidate-lock-v1")
            marker = root / "tracked.txt"
            marker.write_bytes(b"candidate-source-v1")
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            subprocess.run(["git", "-C", str(root), "commit", "--quiet", "-m", "fixture"], check=True)
            normal_target = outer / "normal-target" / "release"
            allocator_target = outer / "allocator-target" / "release"
            normal_target.mkdir(parents=True)
            allocator_target.mkdir(parents=True)
            normal = normal_target / "ppt_lazy_native"
            allocator = allocator_target / "ppt_lazy_native_alloc"
            normal.write_bytes(b"normal-v1")
            allocator.write_bytes(b"allocator-v1")
            os.chmod(normal, 0o755)
            os.chmod(allocator, 0o755)
            manifest_path = outer / "manifest.json"
            manifest_path.write_bytes(b'{"manifest":"v1"}\n')

            def git(*args: str) -> str:
                return subprocess.run(
                    ["git", "-C", str(root), *args], check=True,
                    capture_output=True, text=True,
                ).stdout.strip()

            lock_identity = protocol._identity(lock)
            raw = {
                "root": str(root),
                "ref": git("rev-parse", "HEAD"),
                "tree": git("rev-parse", "HEAD^{tree}"),
                "lock": {
                    "path": "tools/perf-baseline/Cargo.lock",
                    "bytes": lock_identity["bytes"],
                    "sha256": lock_identity["sha256"],
                },
                "normal_binary": protocol._identity(normal),
                "allocator_binary": protocol._identity(allocator),
                "target_dir": str(normal_target),
                "allocator_target_dir": str(allocator_target),
            }
            endpoint = protocol._resolve_endpoint(raw, "candidate", manifest_path, [root])
            self.assertEqual(endpoint["target_dir"], normal_target)

            nested = dict(raw, target_dir=str(root))
            with self.assertRaises(protocol.RunFailure):
                protocol._resolve_endpoint(nested, "candidate", manifest_path, [root])

            lock.write_bytes(b"candidate-lock-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._resolve_endpoint(raw, "candidate", manifest_path, [root])
            lock.write_bytes(b"candidate-lock-v1")

            normal.write_bytes(b"normal-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._resolve_endpoint(raw, "candidate", manifest_path, [root])
            normal.write_bytes(b"normal-v1")

            marker.write_bytes(b"candidate-source-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._resolve_endpoint(raw, "candidate", manifest_path, [root])

    def test_parameterized_manifest_plan_is_build_free(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        self.assertEqual(manifest["baseline"]["ref"], protocol.BASELINE_REF)
        self.assertEqual(manifest["baseline"]["tree"], protocol.BASELINE_TREE)
        self.assertEqual(manifest["baseline"]["lock"]["bytes"], protocol.BASELINE_LOCK_BYTES)
        self.assertEqual(manifest["baseline"]["lock"]["sha256"], protocol.BASELINE_LOCK_SHA256)
        self.assertIsNone(manifest["candidate"])
        self.assertEqual(
            manifest["baseline_adapter"]["compile_fixtures"],
            [
                {"path": path, "bytes": bytes_, "sha256": sha256}
                for path, bytes_, sha256 in protocol.BASELINE_COMPILE_FIXTURES
            ],
        )
        self.assertEqual(set(manifest["build"]["fixtures"]), set(protocol.FIXTURE_KEYS))
        self.assertEqual(manifest["build"]["fixture_routes"], protocol.ROUTE_FIXTURE_KEYS)

        changed = dict(manifest)
        changed["probe_sources"] = list(manifest["probe_sources"])
        changed["probe_sources"][0] = dict(changed["probe_sources"][0], path="tools/perf-baseline/src/unreviewed.rs")
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-probe-allowlist-") as temporary:
            path = Path(temporary) / "manifest.json"
            path.write_bytes(protocol._canonical_json(changed))
            with self.assertRaises(protocol.RunFailure):
                protocol._load_manifest(path, plan=True)
        self.assertEqual(tuple(protocol.ROUTES), (
            "primary_entry_drop",
            "primary_noop_commit",
            "metadata_entry_drop",
            "metadata_noop_commit",
            "animations_entry_drop",
            "animations_noop_commit",
            "broadcast_entry_drop",
            "broadcast_noop_commit",
            "client_data_entry_drop",
            "client_data_noop_commit",
            "diagrams_entry_drop",
            "diagrams_noop_commit",
            "document_comparison_entry_drop",
            "document_comparison_noop_commit",
            "objects_entry_drop",
            "objects_noop_commit",
            "external_media_entry_drop",
            "external_media_noop_commit",
            "fonts_entry_drop",
            "fonts_noop_commit",
            "master_metadata_entry_drop",
            "master_metadata_noop_commit",
            "master_layout_entry_drop",
            "master_layout_noop_commit",
            "picture_bullets_entry_drop",
            "picture_bullets_noop_commit",
            "slide_sync_entry_drop",
            "slide_sync_noop_commit",
        ))

    def test_plan_writes_only_a_small_external_receipt(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-plan-") as temporary:
            output = Path(temporary) / "plan"
            output.mkdir()
            protocol._plan(manifest, MANIFEST, output)
            plan = json.loads((output / "plan.json").read_text(encoding="utf-8"))
            plan_bytes = (output / "plan.json").stat().st_size
        self.assertEqual(plan["status"], "plan_only")
        self.assertEqual(plan["routes"], list(protocol.ROUTES))
        self.assertEqual(plan["abba_order"], list(protocol.ABBA_ORDER))
        self.assertLess(plan_bytes, 1 << 20)

    def test_execution_rejects_unfrozen_candidate(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-manifest-") as temporary:
            path = Path(temporary) / "manifest.json"
            path.write_text(json.dumps(manifest), encoding="utf-8")
            with self.assertRaises(protocol.RunFailure):
                protocol._load_manifest(path, plan=False)

    def test_source_only_protocol_has_fixed_caps(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        protocol_value = manifest["protocol"]
        self.assertEqual(
            manifest["build"]["features"],
            {
                "normal": ["current-api", "current-xlsx-raw-opc"],
                "allocator": ["current-api", "current-xlsx-raw-opc"],
            },
        )
        self.assertEqual(protocol_value["blocks"], 2)
        self.assertEqual(protocol_value["warmups"], 3)
        self.assertGreaterEqual(protocol_value["target_ms"], 500)
        self.assertGreaterEqual(protocol_value["minimum_ms"], 150)
        self.assertEqual(protocol.TARGET_HARD_LIMIT_BYTES, 20_000_000_000)
        self.assertEqual(protocol.MEMORY_HARD_LIMIT_BYTES, 20_000_000_000)
        self.assertEqual(protocol.MAX_CHILD_OUTPUT, 4 * 1024 * 1024)

    def test_strict_json_rejects_duplicate_keys_and_nonfinite_constants(self):
        with self.assertRaises(protocol.RunFailure):
            protocol._strict_json(b'{"route": 1, "route": 2}', "child")
        with self.assertRaises(protocol.RunFailure):
            protocol._strict_json(b'{"value": NaN}', "child")
        with self.assertRaises(protocol.RunFailure):
            protocol._strict_json(b'{"value": 1e999}', "child")
        with self.assertRaises(protocol.RunFailure):
            protocol._finite_number(10**10_000, "child.value")
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-child-") as temporary:
            child = Path(temporary) / "child.json"
            child.write_bytes(b'{"value": 1e999}')
            with self.assertRaises(protocol.RunFailure):
                protocol._parse_child(child, "child")

    def test_candidate_probe_recheck_rejects_source_mutation(self):
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-probe-") as temporary:
            root = Path(temporary) / "candidate"
            root.mkdir()
            subprocess.run(["git", "-C", str(root), "init", "--quiet"], check=True)
            source = root / "probe.rs"
            source.write_bytes(b"before")
            subprocess.run(["git", "-C", str(root), "add", "probe.rs"], check=True)
            digest = hashlib.sha256(source.read_bytes()).hexdigest()
            spec = [{"path": "probe.rs", "bytes": 6, "sha256": digest}]
            endpoint = {"root": root}
            self.assertEqual(protocol._check_probe_sources(spec, endpoint)[0]["sha256"], digest)
            source.write_bytes(b"after!")
            with self.assertRaises(protocol.RunFailure):
                protocol._check_probe_sources(spec, endpoint)

    def test_recheck_binds_actual_manifest_lock_probe_and_binary_bytes(self):
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-recheck-") as temporary:
            outer = Path(temporary)

            def make_endpoint(name):
                root = outer / name
                root.mkdir()
                subprocess.run(["git", "-C", str(root), "init", "--quiet"], check=True)
                subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
                subprocess.run(["git", "-C", str(root), "config", "user.name", "PPT test"], check=True)
                lock = root / "tools" / "perf-baseline" / "Cargo.lock"
                lock.parent.mkdir(parents=True)
                lock.write_bytes(f"{name}-lock-v1".encode())
                probe = root / "probe.rs"
                probe.write_bytes(f"{name}-probe-v1".encode())
                subprocess.run(["git", "-C", str(root), "add", "tools/perf-baseline/Cargo.lock", "probe.rs"], check=True)
                subprocess.run(["git", "-C", str(root), "commit", "--quiet", "-m", "fixture"], check=True)

                target = outer / f"{name}-normal-target" / "release"
                allocator_target = outer / f"{name}-allocator-target" / "release"
                target.mkdir(parents=True)
                allocator_target.mkdir(parents=True)
                normal = target / "ppt_lazy_native"
                allocator = allocator_target / "ppt_lazy_native_alloc"
                normal.write_bytes(f"{name}-normal-binary-v1".encode())
                allocator.write_bytes(f"{name}-allocator-binary-v1".encode())
                os.chmod(normal, 0o755)
                os.chmod(allocator, 0o755)

                def git(*args):
                    return subprocess.run(
                        ["git", "-C", str(root), *args],
                        check=True,
                        capture_output=True,
                        text=True,
                    ).stdout.strip()

                lock_identity = protocol._identity(lock)
                probe_identity = protocol._identity(probe)
                return {
                    "endpoint": {
                        "name": name,
                        "raw": {
                            "lock": {
                                "path": "tools/perf-baseline/Cargo.lock",
                                "bytes": lock_identity["bytes"],
                                "sha256": lock_identity["sha256"],
                            },
                            "normal_binary": protocol._identity(normal),
                            "allocator_binary": protocol._identity(allocator),
                        },
                        "root": root,
                        "head": git("rev-parse", "HEAD"),
                        "tree": git("rev-parse", "HEAD^{tree}"),
                        "target_dir": target.parent,
                        "allocator_target_dir": allocator_target.parent,
                    },
                    "lock": lock,
                    "probe": probe,
                    "probe_spec": {
                        "path": "probe.rs",
                        "bytes": probe_identity["bytes"],
                        "sha256": probe_identity["sha256"],
                    },
                    "normal": normal,
                    "allocator": allocator,
                }

            candidate = make_endpoint("candidate")
            baseline = make_endpoint("baseline")

            manifest_path = outer / "manifest.json"
            manifest_path.write_bytes(b'{"manifest":"v1"}\n')
            projection_path = outer / "projection.json"
            projection_path.write_bytes(b'{"projection":"v1"}\n')
            manifest_identity = protocol._identity(manifest_path)
            protocol._recheck(
                candidate["endpoint"],
                [candidate["probe_spec"]],
                manifest_path,
                manifest_identity,
                "before_candidate",
                projection_path,
                protocol._identity(projection_path),
            )
            protocol._recheck(
                baseline["endpoint"],
                [],
                manifest_path,
                manifest_identity,
                "before_baseline",
            )

            manifest_path.write_bytes(b'{"manifest":"tampered"}\n')
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    candidate["endpoint"],
                    [candidate["probe_spec"]],
                    manifest_path,
                    manifest_identity,
                    "after_manifest_mutation",
                )
            manifest_path.write_bytes(b'{"manifest":"v1"}\n')

            projection_identity = protocol._identity(projection_path)
            projection_path.write_bytes(b'{"projection":"tampered"}\n')
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    candidate["endpoint"],
                    [candidate["probe_spec"]],
                    manifest_path,
                    manifest_identity,
                    "after_projection_mutation",
                    projection_path,
                    projection_identity,
                )
            projection_path.write_bytes(b'{"projection":"v1"}\n')

            candidate["probe"].write_bytes(b"candidate-probe-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    candidate["endpoint"],
                    [candidate["probe_spec"]],
                    manifest_path,
                    manifest_identity,
                    "after_probe_mutation",
                )
            candidate["probe"].write_bytes(b"candidate-probe-v1")

            candidate["lock"].write_bytes(b"candidate-lock-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._check_lock(candidate["endpoint"]["raw"]["lock"], "candidate.lock", candidate["endpoint"]["root"])
            candidate["lock"].write_bytes(b"candidate-lock-v1")

            candidate["normal"].write_bytes(b"candidate-normal-binary-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    candidate["endpoint"],
                    [candidate["probe_spec"]],
                    manifest_path,
                    manifest_identity,
                    "after_candidate_normal_binary_mutation",
                )
            candidate["normal"].write_bytes(b"candidate-normal-binary-v1")

            candidate["allocator"].write_bytes(b"candidate-allocator-binary-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    candidate["endpoint"],
                    [candidate["probe_spec"]],
                    manifest_path,
                    manifest_identity,
                    "after_candidate_allocator_binary_mutation",
                )
            candidate["allocator"].write_bytes(b"candidate-allocator-binary-v1")

            baseline["normal"].write_bytes(b"baseline-normal-binary-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    baseline["endpoint"],
                    [],
                    manifest_path,
                    manifest_identity,
                    "after_baseline_normal_binary_mutation",
                )
            baseline["normal"].write_bytes(b"baseline-normal-binary-v1")

            baseline["allocator"].write_bytes(b"baseline-allocator-binary-tampered")
            with self.assertRaises(protocol.RunFailure):
                protocol._recheck(
                    baseline["endpoint"],
                    [],
                    manifest_path,
                    manifest_identity,
                    "after_baseline_allocator_binary_mutation",
                )

    def test_manifest_rejects_nonfinite_numeric_limits_and_wrong_lock_path(self):
        source = MANIFEST.read_text(encoding="utf-8")
        with tempfile.TemporaryDirectory(prefix="ppt-lazy-invalid-manifest-") as temporary:
            path = Path(temporary) / "manifest.json"
            path.write_text(source.replace('"target_ms": 500', '"target_ms": 1e999'), encoding="utf-8")
            with self.assertRaises(protocol.RunFailure):
                protocol._load_manifest(path, plan=True)
            path.write_text(source.replace('"path": "tools/perf-baseline/Cargo.lock"', '"path": "/tmp/Cargo.lock"'), encoding="utf-8")
            with self.assertRaises(protocol.RunFailure):
                protocol._load_manifest(path, plan=True)

    def test_source_and_resource_overlap_is_rejected_both_directions(self):
        root = Path("/tmp/ppt-lazy-root")
        nested = root / "nested"
        self.assertTrue(protocol._overlap(root, nested))
        self.assertTrue(protocol._overlap(nested, root))

    def test_route_and_variant_identity_mismatches_fail_closed(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        fixture = {"bytes": 1, "sha256": "0" * 64}
        proof = {
            "output": fixture,
            "source_exact": True,
            "changed": False,
            "patch_noop": None,
            "operation_count": None,
            "patch_before_exact": None,
            "patch_after_exact": None,
            "entry": "drop",
            "fixture": {"key": "default", "output": fixture},
        }
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_route_proof(
                dict(proof, output={"bytes": 2, "sha256": "0" * 64}),
                "primary_entry_drop",
                fixture,
                "default",
                "wrong-workload",
            )
        fixtures = {
            key: {"generator": protocol.FIXTURE_GENERATOR, "output": fixture}
            for key in protocol.FIXTURE_KEYS
        }
        report = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "verify",
            "fixtures": fixtures,
            "routes": {"primary_entry_drop": proof},
            "allocation": {"status": "unavailable", "scope": "operation_global_system_allocator"},
        }
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_verify(report, "partial-routes")
        timed = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
            "fixture": {"key": "default", "output": fixture},
            "postcondition": proof,
            "samples": [{
                "elapsed_ns": 150_000_000,
                "iterations": 1,
                "allocation": {
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
                },
            }],
        }
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_timed(
                timed,
                "primary_entry_drop",
                1,
                False,
                manifest["protocol"],
                fixture,
                "default",
                "normal-allocator-mismatch",
            )

    def test_timed_validator_enforces_minimum_sample(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        allocation = {"status": "unavailable", "scope": "operation_global_system_allocator"}
        report_base = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
        }
        short = dict(report_base, samples=[{"elapsed_ns": 149_999_999, "iterations": 1, "allocation": allocation}])
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_timed(
                short,
                "primary_entry_drop",
                1,
                False,
                manifest["protocol"],
                {"bytes": 1, "sha256": "0" * 64},
                "default",
                "short",
            )
        proof = {
            "output": {"bytes": 1, "sha256": "0" * 64},
            "source_exact": True,
            "changed": False,
            "patch_noop": None,
            "operation_count": None,
            "patch_before_exact": None,
            "patch_after_exact": None,
            "entry": "drop",
            "fixture": {"key": "default", "output": {"bytes": 1, "sha256": "0" * 64}},
        }
        valid = dict(report_base, fixture={"key": "default", "output": {"bytes": 1, "sha256": "0" * 64}}, postcondition=proof,
                     samples=[{"elapsed_ns": 150_000_000, "iterations": 1, "allocation": allocation}])
        result = protocol._validate_timed(
            valid,
            "primary_entry_drop",
            1,
            False,
            manifest["protocol"],
            {"bytes": 1, "sha256": "0" * 64},
            "default",
            "valid",
        )
        self.assertEqual(result["samples"][0]["elapsed_ns"], 150_000_000)

    def test_allocator_validator_requires_live_byte_conservation(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            "allocation_calls": 1,
            "deallocation_calls": 0,
            "reallocation_calls": 0,
            "failed_allocation_calls": 0,
            "allocated_bytes": 8,
            "deallocated_bytes": 0,
            "live_bytes_before": 10,
            "live_bytes_after": 10,
            "peak_live_bytes_before": 10,
            "peak_live_bytes_after": 18,
        }
        report = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
            "fixture": {"key": "default", "output": {"bytes": 1, "sha256": "0" * 64}},
            "postcondition": {
                "output": {"bytes": 1, "sha256": "0" * 64},
                "source_exact": True,
                "changed": False,
                "patch_noop": None,
                "operation_count": None,
                "patch_before_exact": None,
                "patch_after_exact": None,
                "entry": "drop",
                "fixture": {"key": "default", "output": {"bytes": 1, "sha256": "0" * 64}},
            },
            "samples": [{"elapsed_ns": 150_000_000, "iterations": 1, "allocation": allocation}],
        }
        with self.assertRaises(protocol.RunFailure):
            protocol._validate_timed(
                report,
                "primary_entry_drop",
                1,
                True,
                manifest["protocol"],
                {"bytes": 1, "sha256": "0" * 64},
                "default",
                "broken-allocation",
            )

    def test_allocator_validator_allows_reallocation_without_allocation_call(self):
        manifest = protocol._load_manifest(MANIFEST, plan=True)
        allocation = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            "allocation_calls": 0,
            "deallocation_calls": 0,
            "reallocation_calls": 1,
            "failed_allocation_calls": 0,
            "allocated_bytes": 0,
            "deallocated_bytes": 0,
            "live_bytes_before": 10,
            "live_bytes_after": 10,
            "peak_live_bytes_before": 10,
            "peak_live_bytes_after": 10,
        }
        digest = {"bytes": 1, "sha256": "0" * 64}
        report = {
            "schema": protocol.PROBE_SCHEMA,
            "mode": "timed",
            "route": "primary_entry_drop",
            "fixture": {"key": "default", "output": digest},
            "postcondition": {
                "output": digest,
                "source_exact": True,
                "changed": False,
                "patch_noop": None,
                "operation_count": None,
                "patch_before_exact": None,
                "patch_after_exact": None,
                "entry": "drop",
                "fixture": {"key": "default", "output": digest},
            },
            "samples": [{"elapsed_ns": 150_000_000, "iterations": 1, "allocation": allocation}],
        }
        protocol._validate_timed(
            report,
            "primary_entry_drop",
            1,
            True,
            manifest["protocol"],
            digest,
            "default",
            "reallocation-only",
        )


if __name__ == "__main__":
    unittest.main()
