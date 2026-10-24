#!/usr/bin/env python3
"""Build-free round-trip and mutation tests for the OLE run validator."""

from __future__ import annotations

import copy
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest import mock

from tools.native_qualification import ole_owner_paired_runner as runner
from tools.native_qualification import prepare_ole_owner_inputs as preparation
from tools.native_qualification import validate_ole_owner_run as validator


def _git(root: Path, *args: str) -> str:
    return subprocess.run(
        ["git", "-C", str(root), *args], check=True, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, text=True,
    ).stdout.strip()


def _source(root: Path, kind: str) -> str:
    root.mkdir()
    _git(root, "init", "--quiet")
    _git(root, "config", "user.email", "validator@test.invalid")
    _git(root, "config", "user.name", "OLE validator")
    (root / "source.txt").write_text(kind, encoding="utf-8")
    cfb = root / preparation.CFB_SOURCE_PATH
    cfb.parent.mkdir(parents=True)
    if kind == "baseline":
        cfb.write_text(
            "pub fn open(reader: R) -> Result<Self, OleError> {}\n"
            "pub fn open_with_limits(mut reader: R, limits: OleFileLimits) {}\n",
            encoding="utf-8",
        )
    else:
        cfb.write_text("pub fn open(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError> {}\n", encoding="utf-8")
    _git(root, "add", ".")
    _git(root, "commit", "--quiet", "-m", kind)
    return _git(root, "rev-parse", "HEAD")


def _sha(path: Path) -> tuple[str, int]:
    data = path.read_bytes()
    return hashlib.sha256(data).hexdigest(), len(data)


def _record(path: Path) -> dict[str, object]:
    digest, size = _sha(path)
    return {"path": str(path), "sha256": digest, "bytes": size}


def _write(path: Path, data: bytes, executable: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    if executable:
        path.chmod(0o755)


class ValidateOleOwnerRunTests(unittest.TestCase):
    def _fixture(self, root: Path) -> tuple[dict[str, Path], Path]:
        baseline = root / "baseline"
        candidate = root / "candidate"
        baseline_ref = _source(baseline, "baseline")
        candidate_ref = _source(candidate, "candidate")
        markers = "\n".join(preparation.REQUIRED_PROBE_MARKERS)
        candidate_probe = candidate / preparation.PROBE_PATH
        candidate_probe.parent.mkdir(parents=True)
        candidate_probe.write_text(
            markers + "\nOleFile::open(\n        Cursor::new(package.into_inner()),\n        OleFileLimits::default(),\nOleFileLimits::default()\n",
            encoding="utf-8",
        )
        _git(candidate, "add", ".")
        _git(candidate, "commit", "--quiet", "-m", "candidate probe")
        candidate_ref = _git(candidate, "rev-parse", "HEAD")
        reviewed = root / "reviewed.rs"
        reviewed.write_text(markers + "\nOleFile::open(Cursor::new(package.into_inner()))\n", encoding="utf-8")
        clone = root / "clone.rs"
        clone.write_text("\n".join(preparation.REQUIRED_CLONE_MARKERS), encoding="utf-8")
        oracle = root / "oracle.rs"
        oracle.write_text("fn check_ole_outputs\nallocator_evidence\nowner_target\n", encoding="utf-8")
        lock = root / "Cargo.lock"
        lock.write_text("fixture lock\n", encoding="utf-8")
        pins = {
            "BASELINE_COMMIT": baseline_ref,
            "REVIEWED_PROBE_SHA256": _sha(reviewed)[0],
            "CANDIDATE_ADAPTER_SHA256": _sha(candidate_probe)[0],
            "CLONE_PROBE_SHA256": _sha(clone)[0],
            "ORACLE_SOURCE_SHA256": _sha(oracle)[0],
            "LOCK_SHA256": _sha(lock)[0],
        }
        prep_path = root / "preparation.json"
        with mock.patch.multiple(preparation, **pins):
            preparation.prepare(
                baseline_root=baseline,
                candidate_root=candidate,
                candidate_ref=candidate_ref,
                reviewed_probe=reviewed,
                clone_probe=clone,
                oracle_probe=oracle,
                lock=lock,
                drivers_root=root / "drivers",
                output=prep_path,
            )
        manifest = json.loads((Path(__file__).parents[0] / "native_qualification/manifests/ole-owner.json").read_text())
        manifest["input_pins"].update({
            "reviewed_probe_sha256": pins["REVIEWED_PROBE_SHA256"],
            "candidate_adapter_sha256": pins["CANDIDATE_ADAPTER_SHA256"],
            "clone_probe_sha256": pins["CLONE_PROBE_SHA256"],
            "oracle_source_sha256": pins["ORACLE_SOURCE_SHA256"],
            "lock_sha256": pins["LOCK_SHA256"],
        })
        manifest_path = root / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, sort_keys=True), encoding="utf-8")
        build_dir = root / "run"
        target_root = build_dir / "targets"
        build_dir.mkdir()

        identities: dict[str, dict[str, object]] = {}
        for role, endpoint in (("before", "baseline"), ("after", "candidate")):
            source = json.loads(prep_path.read_text())[endpoint]
            source_keys = ("root", "head", "tree", "status", "status_paths", "clean", "status_sha256", "diff_sha256", "production_tree_sha256")
            identity: dict[str, object] = {
                "source": {key: source[key] for key in source_keys},
                "driver_root": source["driver"]["root"],
                "driver_files": source["driver"]["files"],
            }
            identity["identity_sha256"] = hashlib.sha256(json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
            identities[role] = identity

        build_events = []
        builds: dict[str, object] = {}
        for key in runner.BUILD_KEYS:
            role, variant = key.split("/")
            target = target_root / f"{role}-{variant}"
            binaries = {
                "probe": target / "release/examples/ole_w5_probe",
                "clone": target / "release/examples/ole_package_clone_probe",
                "oracle": target / "release/examples/ole_owner_oracle",
            }
            artifacts = {}
            outputs = []
            for name, path in binaries.items():
                _write(path, f"{key}-{name}".encode(), executable=True)
                artifacts[name] = _record(path)
                outputs.append({**_record(path), "exists": True})
            stdout = build_dir / "commands" / f"{key.replace('/', '-')}.stdout"
            stderr = build_dir / "commands" / f"{key.replace('/', '-')}.stderr"
            _write(stdout, b"")
            _write(stderr, b"")
            event = {
                "event_id": f"build-{key.replace('/', '-')}", "phase": "build", "role": role, "variant": variant,
                "source_identity_sha256": identities[role]["identity_sha256"],
                "source_identity_before": identities[role], "source_identity_after": identities[role],
                "returncode": 0, "runner_error": None, "process_guard_violations": [],
                "post_exit_descendants": [], "post_exit_offenders": [],
                "environment": {"CARGO_TARGET_DIR": str(target), "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0", "LITCHI_MEM_MAX": "20G"},
                "stdout": _record(stdout), "stderr": _record(stderr), "outputs": outputs,
            }
            build_events.append(event)
            builds[key] = {"role": role, "variant": variant, "features": validator.EXPECTED_FEATURES[key], "target": str(target), "command": [], "event_id": event["event_id"], "artifacts": artifacts}
        build_events_path = build_dir / "events-build.jsonl"
        build_events_path.write_text("\n".join(json.dumps(event, sort_keys=True) for event in build_events) + "\n", encoding="utf-8")
        event_raw = build_events_path.read_bytes()
        build_receipt = {
            "schema": runner.SCHEMA, "phase": "build", "status": "passed",
            "manifest": {"path": str(manifest_path), "sha256": _sha(manifest_path)[0]},
            "preparation": {"path": str(prep_path), "sha256": _sha(prep_path)[0]},
            "runner": _record(Path(runner.__file__)), "limits": {"target_preflight_bytes": validator.TARGET_PREFLIGHT_BYTES, "target_hard_bytes": validator.TARGET_HARD_BYTES, "memory_bytes": validator.MEMORY_BYTES},
            "claim_scope": validator.CLAIM_FREE["build"], "target_root": str(target_root),
            "event_journal": {"path": str(build_events_path), "count": len(build_events), "event_ids": [e["event_id"] for e in build_events], "sha256": hashlib.sha256(json.dumps(build_events, sort_keys=True, separators=(",", ":")).encode()).hexdigest(), "file_sha256": hashlib.sha256(event_raw).hexdigest(), "bytes": len(event_raw)},
            "toolchain": {name: {**_record(Path(os.sys.executable)), "version": "fixture"} for name in ("cargo", "rustc")},
            "identities": identities, "builds": builds,
        }
        build_path = build_dir / "build-receipt.json"
        build_path.write_text(json.dumps(build_receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")

        reports = {}
        metadata = "\n".join(["schema\t1", "owner_streams\t65", "owner_target\tole_w5_edit_target.bin", "owner_original_bytes\t15", "owner_replacement_bytes\t19", "properties\tvalidated", "doc\trevision\tInsertion\t0\t24\tW5Oracle", "xls\tCells\t0\t0\t4242"]) + "\n"
        allocator_header = "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches"
        allocator_rows = [allocator_header]
        for route in validator.ALLOCATOR_ROUTES:
            retained = "129\t128" if route in {"doc_history", "xls_history"} else "0\t0"
            allocator_rows.append(f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{retained}")
        for key in runner.BUILD_KEYS:
            role, variant = key.split("/")
            directory = build_dir / "oracle" / key.replace("/", "-")
            files = {}
            values = {
                "ole-source.bin": b"source", "ole-noop.bin": b"source", "ole-replacement.bin": b"replacement",
                "properties.tsv": b"link_base\thttps://example.invalid/base/\nhyperlinks\t1\nhyperlink\t7\t11\t0\thttps://example.invalid/target\tbookmark\nopaque_property\t42\t128\n", "doc-source.bin": b"doc-source", "doc-edit.bin": b"doc-edit",
                "xls-source.bin": b"xls-source", "xls-edit.bin": b"xls-edit", "oracle.tsv": metadata.encode(),
                "allocator.tsv": (b"status\tnative-uninstrumented\n" if variant == "normal" else ("\n".join(allocator_rows) + "\n").encode()),
            }
            for name, data in values.items():
                path = directory / name
                _write(path, data)
                files[name] = _record(path)
            reports[key] = {"role": role, "variant": variant, "directory": str(directory), "files": files, "event_id": f"oracle-{key.replace('/', '-') }"}

        timing = {"clone": [{"block": block, "legs": [{"role": role, "rows": [{"nanoseconds": validator.MIN_SPAN_NS, "output_bytes": 1} for _ in range(2 * validator.SAMPLES)]} for role in validator.ABBA_ORDER]} for block in range(validator.BLOCKS)], "routes": {route: [{"block": block, "legs": [{"role": role, "rows": [{"nanoseconds": validator.MIN_SPAN_NS, "output_bytes": 1} for _ in range(validator.SAMPLES)]} for role in validator.ABBA_ORDER]} for block in range(validator.BLOCKS)] for route in validator.ROUTES}}
        allocation = {"abba_order": list(validator.ABBA_ORDER), "elapsed_claimable": False, "routes": {route: {"legs": [{"role": role, "report": {"output_bytes": 1}} for role in validator.ABBA_ORDER]} for route in validator.ALLOCATOR_ROUTES}}
        measurements = {"schema": validator.MEASUREMENTS_SCHEMA, "status": "complete; coordinator review pending", "routes": list(validator.ROUTES), "abba_order": list(validator.ABBA_ORDER), "blocks": validator.BLOCKS, "warmup": validator.WARMUP, "samples": validator.SAMPLES, "minimum_span_ns": validator.MIN_SPAN_NS, "timing": timing, "allocation": allocation, "oracle": {"status": "preflight-validated", "files": reports, "claim_scope": validator.CLAIM_FREE["oracle"]}, "claim_scope": validator.CLAIM_FREE["measurements"]}
        measurements_path = build_dir / "measurements.json"
        measurements_path.write_text(json.dumps(measurements, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        abba_stdout = build_dir / "runs/abba.stdout"
        abba_stderr = build_dir / "runs/abba.stderr"
        _write(abba_stdout, b"ok\n")
        _write(abba_stderr, b"")
        abba_event = {"event_id": "abba-fixture", "phase": "abba", "returncode": 0, "runner_error": None, "process_guard_violations": [], "post_exit_descendants": [], "post_exit_offenders": [], "stdout": _record(abba_stdout), "stderr": _record(abba_stderr)}
        abba_events = build_dir / "events-abba.jsonl"
        abba_events.write_text(json.dumps(abba_event, sort_keys=True) + "\n", encoding="utf-8")
        abba_raw = abba_events.read_bytes()
        abba_receipt = {"schema": runner.SCHEMA, "phase": "abba", "status": "passed", "manifest": {"path": str(manifest_path), "sha256": _sha(manifest_path)[0]}, "preparation": {"path": str(prep_path), "sha256": _sha(prep_path)[0]}, "runner": _record(Path(runner.__file__)), "claim_scope": "orchestration and source-bound evidence only; no performance acceptance claim", "build_receipt_sha256": _sha(build_path)[0], "event_journal": {"path": str(abba_events), "count": 1, "event_ids": ["abba-fixture"], "sha256": hashlib.sha256(json.dumps([abba_event], sort_keys=True, separators=(",", ":")).encode()).hexdigest(), "file_sha256": hashlib.sha256(abba_raw).hexdigest(), "bytes": len(abba_raw)}, "measurements": {"path": str(measurements_path), "sha256": _sha(measurements_path)[0], "bytes": _sha(measurements_path)[1]}}
        abba_path = build_dir / "abba-receipt.json"
        abba_path.write_text(json.dumps(abba_receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return {"manifest": manifest_path, "preparation": prep_path, "build": build_path, "abba": abba_path, "measurements": measurements_path}, runner.Path(runner.__file__) if False else Path(runner.__file__)

    def test_build_free_preparation_to_validator_round_trip(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            result = validator.validate(
                manifest_path=paths["manifest"], preparation_path=paths["preparation"], build_path=paths["build"],
                abba_path=paths["abba"], measurements_path=paths["measurements"], runner_path=runner_path,
            )
            self.assertEqual(result["status"], "validated")

    def test_rejects_mutated_immutable_build_receipt_hash(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            receipt = json.loads(paths["abba"].read_text())
            receipt["build_receipt_sha256"] = "0" * 64
            paths["abba"].write_text(json.dumps(receipt))
            with self.assertRaisesRegex(validator.ValidationError, "build receipt hash"):
                validator.validate(
                    manifest_path=paths["manifest"], preparation_path=paths["preparation"], build_path=paths["build"],
                    abba_path=paths["abba"], measurements_path=paths["measurements"], runner_path=runner_path,
                )

    def test_rejects_build_feature_or_log_mutation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            receipt = json.loads(paths["build"].read_text())
            receipt["builds"]["after/normal"]["features"] = "legacy-owner-api"
            paths["build"].write_text(json.dumps(receipt))
            with self.assertRaisesRegex(validator.ValidationError, "features"):
                validator.validate(
                    manifest_path=paths["manifest"], preparation_path=paths["preparation"], build_path=paths["build"],
                    abba_path=paths["abba"], measurements_path=paths["measurements"], runner_path=runner_path,
                )

        with tempfile.TemporaryDirectory(prefix="ole-owner-validator-") as directory:
            paths, runner_path = self._fixture(Path(directory))
            log = Path(directory) / "run/commands/before-normal.stdout"
            log.write_bytes(b"tampered\n")
            with self.assertRaisesRegex(validator.ValidationError, "stdout SHA-256"):
                validator.validate(
                    manifest_path=paths["manifest"], preparation_path=paths["preparation"], build_path=paths["build"],
                    abba_path=paths["abba"], measurements_path=paths["measurements"], runner_path=runner_path,
                )

    def test_allocator_conservation_rejects_counter_mutation(self) -> None:
        header = "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches"
        rows = [
            "read\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1080\t110\t1100\t0\t0",
        ]
        for route in validator.ALLOCATOR_ROUTES[1:]:
            retained = "129\t128" if route in {"doc_history", "xls_history"} else "0\t0"
            rows.append(f"{route}\t10\t100\t2\t20\t1\t5\t6\t100\t1000\t108\t1081\t110\t1100\t{retained}")
        with self.assertRaisesRegex(validator.ValidationError, "live-byte conservation"):
            validator._validate_allocator_rows([header, *rows], "fixture")


if __name__ == "__main__":
    unittest.main()
