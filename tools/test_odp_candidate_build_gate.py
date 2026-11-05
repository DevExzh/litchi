#!/usr/bin/env python3
"""Build-free contract tests for the ODP candidate source-build handoff."""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock


MODULE_PATH = Path(__file__).with_name("native_qualification") / "odp_candidate_build_gate.py"
SPEC = importlib.util.spec_from_file_location("odp_candidate_build_gate", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
GATE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = GATE
SPEC.loader.exec_module(GATE)


def _digest(path: Path) -> dict[str, object]:
    value = path.read_bytes()
    return {"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}


def _build_spec() -> dict[str, object]:
    return {
        "runner": GATE.BUILD_RUNNER,
        "argv_template": [
            "cargo", "build", "--offline", "--locked", "--jobs", "1", "--release",
            "--manifest-path", "{candidate_root}/tools/perf-baseline/Cargo.toml",
            "--target-dir", "{target_dir}", "--bin", "odp_native_qualification_probe",
        ],
        "environment": {
            "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_RELEASE_DEBUG": "0", "CARGO_TARGET_DIR": "{target_dir}",
            "LITCHI_MEM_MAX": "20G", "TMPDIR": "{tmpdir}",
        },
        "claim_scope": GATE.BUILD_CLAIM,
    }


class OdpCandidateBuildGateTests(unittest.TestCase):
    def test_legacy_unstructured_cargo_string_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "binder.json"
            path.write_text(json.dumps({
                "schema": "litchi.odp.native-qualification.c2a-source-binder.v1",
                "candidate": {"commit": GATE.CANDIDATE_REF, "tree": GATE.CANDIDATE_TREE},
                "build_command": "cargo build --offline --locked",
            }), encoding="utf-8")
            with self.assertRaises(GATE.RunFailure):
                GATE._load_binder(path)

    def test_structured_contract_contains_target_lock_and_capped_environment(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "binder.json"
            path.write_text(json.dumps({
                "schema": "litchi.odp.native-qualification.c2a-source-binder.v1",
                "candidate": {"commit": GATE.CANDIDATE_REF, "tree": GATE.CANDIDATE_TREE},
                "build": _build_spec(),
            }), encoding="utf-8")
            binder = GATE._load_binder(path)
            spec = binder["build"]
            self.assertIn("--offline", spec["argv_template"])
            self.assertIn("--locked", spec["argv_template"])
            self.assertIn("--target-dir", spec["argv_template"])
            self.assertEqual(spec["environment"]["CARGO_BUILD_JOBS"], "1")
            self.assertEqual(spec["environment"]["LITCHI_MEM_MAX"], "20G")

    def test_public_prepare_records_actual_targeted_argv_without_building(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "candidate"
            root.mkdir()
            (root / ".cargo").mkdir()
            files = {
                "Cargo.toml": "[workspace]\nmembers = []\n",
                "Cargo.lock": "# lock\n",
                "rust-toolchain.toml": "[toolchain]\nchannel = \"stable\"\n",
                ".cargo/config.toml": "[build]\njobs = 1\n",
                "tools/perf-baseline/Cargo.toml": "[package]\nname = \"perf\"\nversion = \"0.0.0\"\n",
                "tools/perf-baseline/Cargo.lock": "# perf lock\n",
                "tools/perf-baseline/src/bin/odp_native_qualification_probe.rs": "fn main() {}\n",
                "tools/odp_native_qualification.py": "# reviewed runner\n",
            }
            for relative, content in files.items():
                path = root / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(content, encoding="utf-8")
            fixture = Path(directory) / "fixture.odp"
            fixture.write_bytes(b"controlled fixture")
            subprocess.run(["git", "-C", str(root), "init", "-q"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.email", "test@example.invalid"], check=True)
            subprocess.run(["git", "-C", str(root), "config", "user.name", "Test"], check=True)
            subprocess.run(["git", "-C", str(root), "add", "."], check=True)
            subprocess.run(["git", "-C", str(root), "commit", "-q", "-m", "candidate"], check=True)
            commit = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
            tree = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD^{tree}"], text=True).strip()
            binder = Path(directory) / "binder.json"
            binder.write_text(json.dumps({
                "schema": "litchi.odp.native-qualification.c2a-source-binder.v1",
                "candidate": {
                    "root": str(root), "commit": commit, "tree": tree,
                    "cargo_lock": {"path": "Cargo.lock", **_digest(root / "Cargo.lock")},
                    "probe": {"path": GATE.PROBE_RELATIVE, **_digest(root / GATE.PROBE_RELATIVE)},
                },
                "fixture": {"path": str(fixture), **_digest(fixture), "provenance": "controlled"},
                "build": _build_spec(),
            }), encoding="utf-8")
            output = Path(directory) / "preparation.json"
            target = Path(directory) / "target"
            build_output = Path(directory) / "build-output"
            tmpdir = Path(directory) / "tmp"
            tmpdir.mkdir()
            with mock.patch.object(GATE, "CANDIDATE_REF", commit), mock.patch.object(GATE, "CANDIDATE_TREE", tree):
                plan = GATE.prepare(
                    binder, target_dir=target, build_output=build_output,
                    tmpdir=tmpdir, cpu=3, output=output,
                )
            self.assertEqual(plan["claim_scope"], "source preparation only; no build, correctness, timing, or acceptance claim")
            argv = plan["command"]["argv"]
            self.assertEqual(argv[0:6], ["cargo", "build", "--offline", "--locked", "--jobs", "1"])
            self.assertIn("--target-dir", argv)
            self.assertEqual(argv[argv.index("--target-dir") + 1], str(target))
            self.assertFalse(target.exists())
            self.assertFalse(build_output.exists())
            with mock.patch.object(GATE, "CANDIDATE_REF", commit), mock.patch.object(GATE, "CANDIDATE_TREE", tree):
                self.assertEqual(GATE._load_preparation(output)["status"], "source_prepared")
                tampered = json.loads(output.read_text(encoding="utf-8"))
                tampered["command"]["argv"] = [item for item in tampered["command"]["argv"] if item != "--locked"]
                output.write_text(json.dumps(tampered), encoding="utf-8")
                with self.assertRaises(GATE.RunFailure):
                    GATE._load_preparation(output)

    def test_build_passes_the_four_megabyte_output_cap_to_stream_guard(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source"
            source.mkdir()
            preparation_path = root / "preparation.json"
            preparation_path.write_text("{}", encoding="utf-8")
            target = root / "target"
            build_output = root / "build-output"
            artifact = target / "release" / "odp_native_qualification_probe"
            preparation = {
                "candidate": {"root": str(source), "files": {}},
                "target_dir": str(target), "build_output": str(build_output),
                "tmpdir": str(root / "tmp"), "cpu": 0, "timeout_s": 1, "max_load": 1000.0,
                "command": {"argv": ["cargo", "build"], "environment": {"variables": {}}},
                "artifact": {"path": str(artifact)}, "fixture": {},
            }
            guard = {"phase": "guard", "load_1m": 0.0, "max_load": 1000.0, "heavy_processes": [], "checked_unix_ns": 1}

            def stream(*args: object, **kwargs: object) -> dict[str, object]:
                target_path = args[4]
                assert isinstance(target_path, Path)
                artifact_path = target_path / "release" / "odp_native_qualification_probe"
                artifact_path.parent.mkdir(parents=True)
                artifact_path.write_bytes(b"#!/bin/sh\nexit 0\n")
                artifact_path.chmod(0o755)
                return {"label": "candidate-normal"}

            with mock.patch.object(GATE, "_load_preparation", return_value=preparation), \
                 mock.patch.object(GATE, "_source_state", side_effect=lambda *args, **kwargs: {"root": str(source)}), \
                 mock.patch.object(GATE, "_quiet_guard", return_value=guard), \
                 mock.patch.object(GATE, "_stream_command", side_effect=stream) as streamed:
                GATE.build(preparation_path)
            self.assertEqual(streamed.call_args.kwargs["max_output_bytes"], GATE.MAX_OUTPUT_BYTES)


if __name__ == "__main__":
    unittest.main()
