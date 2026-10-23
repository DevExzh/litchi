#!/usr/bin/env python3
"""Fail-closed provenance tests for the native XLS F3 runner."""

from __future__ import annotations

import importlib.util
import subprocess
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
RUNNER_PATH = ROOT / "tools" / "xls_f3_qualification_abba.py"
SPEC = importlib.util.spec_from_file_location("xls_f3_runner_under_test", RUNNER_PATH)
assert SPEC is not None and SPEC.loader is not None
RUNNER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(RUNNER)


def git(directory: Path, *arguments: str) -> str:
    completed = subprocess.run(
        ["git", "-C", str(directory), *arguments],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return completed.stdout.strip()


def initialize_repository(directory: Path, files: dict[str, bytes]) -> str:
    directory.mkdir(parents=True)
    git(directory, "init", "--initial-branch=main")
    git(directory, "config", "user.email", "xls-f3-test@example.invalid")
    git(directory, "config", "user.name", "XLS F3 provenance test")
    for relative, contents in files.items():
        path = directory / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(contents)
    git(directory, "add", "--all")
    git(directory, "commit", "-m", "fixture")
    return git(directory, "rev-parse", "HEAD")


class ProvenanceTests(unittest.TestCase):
    def test_status_paths_preserves_porcelain_prefixes(self) -> None:
        self.assertEqual(
            RUNNER.status_paths(" M tools/perf-baseline/src/lib.rs\n?? tools/new.rs\n"),
            ["tools/perf-baseline/src/lib.rs", "tools/new.rs"],
        )

    def test_only_frozen_overlay_is_admitted_and_production_mutation_is_rejected(self) -> None:
        # The checkout is read-only input: all writes below target the temporary
        # source fixture, while frozen probe bytes come from `git show`.
        expected_probe = RUNNER.probe_identities(ROOT, RUNNER.COMMON_HARNESS_REF)
        with tempfile.TemporaryDirectory(prefix="xls-f3-provenance-") as name:
            source = Path(name) / "source"
            source_ref = initialize_repository(source, {"production.txt": b"baseline"})
            for relative in RUNNER.PROBE_PATHS:
                path = source / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(
                    RUNNER.git_show_bytes(ROOT, RUNNER.COMMON_HARNESS_REF, relative)
                )
            expected_production = RUNNER.hash_git_manifest(
                source, source_ref, RUNNER.PROBE_PATHS
            )
            expected_source = RUNNER.source_manifest_sha256(
                expected_production, expected_probe
            )
            identity = RUNNER.source_identity(
                source,
                source_ref=source_ref,
                harness_ref=RUNNER.COMMON_HARNESS_REF,
                expected_probe=expected_probe,
                expected_production_manifest=expected_production,
                expected_source_manifest=expected_source,
            )
            self.assertEqual(identity["production_manifest_sha256"], expected_production)
            self.assertEqual(identity["source_manifest_sha256"], expected_source)
            self.assertEqual(identity["overlay_paths"], list(RUNNER.PROBE_PATHS))

            (source / "production.txt").write_bytes(b"tampered")
            with self.assertRaises(RUNNER.RunnerError):
                RUNNER.source_identity(
                    source,
                    source_ref=source_ref,
                    harness_ref=RUNNER.COMMON_HARNESS_REF,
                    expected_probe=expected_probe,
                    expected_production_manifest=expected_production,
                    expected_source_manifest=expected_source,
                )

    def test_unexpected_overlay_path_is_rejected(self) -> None:
        expected_probe = RUNNER.probe_identities(ROOT, RUNNER.COMMON_HARNESS_REF)
        with tempfile.TemporaryDirectory(prefix="xls-f3-provenance-") as name:
            source = Path(name) / "source"
            source_ref = initialize_repository(source, {"production.txt": b"baseline"})
            for relative in RUNNER.PROBE_PATHS:
                path = source / relative
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(
                    RUNNER.git_show_bytes(ROOT, RUNNER.COMMON_HARNESS_REF, relative)
                )
            expected_production = RUNNER.hash_git_manifest(
                source, source_ref, RUNNER.PROBE_PATHS
            )
            expected_source = RUNNER.source_manifest_sha256(
                expected_production, expected_probe
            )
            (source / "unexpected.txt").write_bytes(b"not an overlay")
            with self.assertRaises(RUNNER.RunnerError):
                RUNNER.source_identity(
                    source,
                    source_ref=source_ref,
                    harness_ref=RUNNER.COMMON_HARNESS_REF,
                    expected_probe=expected_probe,
                    expected_production_manifest=expected_production,
                    expected_source_manifest=expected_source,
                )


if __name__ == "__main__":
    unittest.main()
