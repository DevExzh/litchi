#!/usr/bin/env python3
"""Source-only tests for the PPT lazy-native baseline preparation."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from tools.native_qualification import prepare_ppt_lazy_native_inputs as prep


def git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    return result.stdout.strip()


def repository(
    root: Path,
    *,
    baseline: bool,
    fixture: bytes | None = None,
    compile_fixtures: dict[Path, bytes] | None = None,
) -> str:
    root.mkdir(parents=True)
    git(root, "init", "-q", "--initial-branch=main")
    git(root, "config", "user.email", "ppt-prep@example.invalid")
    git(root, "config", "user.name", "PPT preparation test")
    (root / "tools" / "perf-baseline" / "src" / "bin").mkdir(parents=True)
    (root / "tools" / "perf-baseline" / "Cargo.lock").write_bytes(
        b"baseline lock\n" if baseline else b"candidate lock\n"
    )
    lib = "pub mod pptx_retention;\n"
    if not baseline:
        lib += f"{prep.REGISTRATION_LINE}\n"
    (root / "tools" / "perf-baseline" / "src" / "lib.rs").write_text(lib, encoding="utf-8")
    if not baseline:
        for relative in prep.PROBE_PATHS:
            (root / relative).write_text(f"{relative}\n", encoding="utf-8")
    if baseline:
        (root / ".gitignore").write_text("/test-data/compile-fixture-1.bin\n", encoding="utf-8")
    if fixture is not None:
        destination = root / prep.FIXTURE_PATH
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(fixture)
    for relative, value in (compile_fixtures or {}).items():
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(value)
    git(root, "add", ".")
    git(root, "commit", "-qm", "fixture")
    return git(root, "rev-parse", "HEAD")


class PptPreparationTests(unittest.TestCase):
    def test_materializes_and_binds_the_reviewed_fixture(self) -> None:
        fixture = b"reviewed fixture\n"
        compile_fixture_values = {
            Path(f"test-data/compile-fixture-{index}.bin"): f"compile fixture {index}\n".encode()
            for index in range(2)
        }
        compile_fixtures = tuple(
            (path, len(value), hashlib.sha256(value).hexdigest())
            for path, value in compile_fixture_values.items()
        )
        with tempfile.TemporaryDirectory(prefix="ppt-prep-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            old_values = (
                prep.BASELINE_REF,
                prep.BASELINE_TREE,
                prep.BASELINE_LOCK_BYTES,
                prep.BASELINE_LOCK_SHA256,
                prep.FIXTURE_BYTES,
                prep.FIXTURE_SHA256,
                prep.COMPILE_FIXTURES,
            )
            prep.COMPILE_FIXTURES = compile_fixtures
            baseline_ref = repository(baseline, baseline=True)
            candidate_ref = repository(
                candidate,
                baseline=False,
                fixture=fixture,
                compile_fixtures=compile_fixture_values,
            )
            baseline_tree = git(baseline, "rev-parse", "HEAD^{tree}")
            prep.BASELINE_REF = baseline_ref
            prep.BASELINE_TREE = baseline_tree
            prep.BASELINE_LOCK_BYTES = len(b"baseline lock\n")
            prep.BASELINE_LOCK_SHA256 = hashlib.sha256(b"baseline lock\n").hexdigest()
            prep.FIXTURE_BYTES = len(fixture)
            prep.FIXTURE_SHA256 = hashlib.sha256(fixture).hexdigest()
            output = root / "preparation.json"
            try:
                receipt = prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    output=output,
                )
            finally:
                (
                    prep.BASELINE_REF,
                    prep.BASELINE_TREE,
                    prep.BASELINE_LOCK_BYTES,
                    prep.BASELINE_LOCK_SHA256,
                    prep.FIXTURE_BYTES,
                    prep.FIXTURE_SHA256,
                    prep.COMPILE_FIXTURES,
                ) = old_values
            materialized = baseline / prep.FIXTURE_PATH
            self.assertEqual(materialized.read_bytes(), fixture)
            self.assertEqual(receipt["baseline"]["fixture"]["destination"]["bytes"], len(fixture))
            self.assertEqual(
                receipt["baseline"]["fixture"]["destination"]["sha256"],
                hashlib.sha256(fixture).hexdigest(),
            )
            self.assertEqual(len(receipt["baseline"]["compile_fixtures"]), len(compile_fixtures))
            for index, (path, value) in enumerate(compile_fixture_values.items()):
                overlay = receipt["baseline"]["compile_fixtures"][index]
                self.assertEqual(overlay["source"]["path"], str(candidate / path))
                self.assertEqual(overlay["destination"]["path"], str(baseline / path))
                self.assertEqual(overlay["destination"]["bytes"], len(value))
                self.assertEqual(overlay["destination"]["sha256"], hashlib.sha256(value).hexdigest())
            self.assertEqual(receipt["baseline"]["tree"], baseline_tree)
            self.assertEqual(
                json.loads(output.read_text(encoding="utf-8"))["claim_scope"],
                "source preparation identity only; no build, correctness, timing, or acceptance claim",
            )
            self.assertEqual(
                sorted(prep._status_paths(receipt["baseline"]["after_overlay"]["status"])),
                sorted(
                    [str(path) for path in prep.PROBE_PATHS]
                    + [str(prep.REGISTRATION_PATH), str(prep.FIXTURE_PATH)]
                    + [str(path) for path, _, _ in compile_fixtures]
                ),
            )

    def test_rejects_a_changed_reviewed_fixture(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-prep-fixture-guard-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            old_values = (
                prep.BASELINE_REF,
                prep.BASELINE_TREE,
                prep.BASELINE_LOCK_BYTES,
                prep.BASELINE_LOCK_SHA256,
                prep.FIXTURE_BYTES,
                prep.FIXTURE_SHA256,
                prep.COMPILE_FIXTURES,
            )
            prep.COMPILE_FIXTURES = ()
            baseline_ref = repository(baseline, baseline=True)
            candidate_ref = repository(candidate, baseline=False, fixture=b"fixture\n")
            baseline_tree = git(baseline, "rev-parse", "HEAD^{tree}")
            prep.BASELINE_REF = baseline_ref
            prep.BASELINE_TREE = baseline_tree
            prep.BASELINE_LOCK_BYTES = len(b"baseline lock\n")
            prep.BASELINE_LOCK_SHA256 = hashlib.sha256(b"baseline lock\n").hexdigest()
            prep.FIXTURE_BYTES = 8
            prep.FIXTURE_SHA256 = hashlib.sha256(b"wrong\n").hexdigest()
            try:
                with self.assertRaisesRegex(prep.PreparationError, "reviewed fixture"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                        output=root / "missing.json",
                    )
            finally:
                (
                    prep.BASELINE_REF,
                    prep.BASELINE_TREE,
                    prep.BASELINE_LOCK_BYTES,
                    prep.BASELINE_LOCK_SHA256,
                    prep.FIXTURE_BYTES,
                    prep.FIXTURE_SHA256,
                    prep.COMPILE_FIXTURES,
                ) = old_values

    def test_materializes_a_sparse_tracked_compile_fixture(self) -> None:
        fixture = b"reviewed fixture\n"
        compile_path = Path("test-data/compile-fixture-tracked.bin")
        compile_value = b"tracked compile fixture\n"
        compile_fixtures = (
            (compile_path, len(compile_value), hashlib.sha256(compile_value).hexdigest()),
        )
        with tempfile.TemporaryDirectory(prefix="ppt-prep-sparse-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            old_values = (
                prep.BASELINE_REF,
                prep.BASELINE_TREE,
                prep.BASELINE_LOCK_BYTES,
                prep.BASELINE_LOCK_SHA256,
                prep.FIXTURE_BYTES,
                prep.FIXTURE_SHA256,
                prep.COMPILE_FIXTURES,
            )
            prep.COMPILE_FIXTURES = compile_fixtures
            try:
                baseline_ref = repository(
                    baseline,
                    baseline=True,
                    compile_fixtures={compile_path: compile_value},
                )
                candidate_ref = repository(
                    candidate,
                    baseline=False,
                    fixture=fixture,
                    compile_fixtures={compile_path: compile_value},
                )
                baseline_tree = git(baseline, "rev-parse", "HEAD^{tree}")
                subprocess.run(
                    ["git", "-C", str(baseline), "update-index", "--skip-worktree", "--", str(compile_path)],
                    check=True,
                )
                (baseline / compile_path).unlink()
                self.assertEqual(git(baseline, "status", "--porcelain=v1"), "")
                prep.BASELINE_REF = baseline_ref
                prep.BASELINE_TREE = baseline_tree
                prep.BASELINE_LOCK_BYTES = len(b"baseline lock\n")
                prep.BASELINE_LOCK_SHA256 = hashlib.sha256(b"baseline lock\n").hexdigest()
                prep.FIXTURE_BYTES = len(fixture)
                prep.FIXTURE_SHA256 = hashlib.sha256(fixture).hexdigest()
                receipt = prep.prepare(
                    baseline_root=baseline,
                    candidate_root=candidate,
                    candidate_ref=candidate_ref,
                    candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                    output=root / "preparation.json",
                )
            finally:
                (
                    prep.BASELINE_REF,
                    prep.BASELINE_TREE,
                    prep.BASELINE_LOCK_BYTES,
                    prep.BASELINE_LOCK_SHA256,
                    prep.FIXTURE_BYTES,
                    prep.FIXTURE_SHA256,
                    prep.COMPILE_FIXTURES,
                ) = old_values
            self.assertEqual((baseline / compile_path).read_bytes(), compile_value)
            self.assertIn(str(compile_path), receipt["baseline"]["overlay_paths"])
            self.assertEqual(
                receipt["baseline"]["compile_fixtures"][0]["destination"]["sha256"],
                hashlib.sha256(compile_value).hexdigest(),
            )
            self.assertNotIn(
                str(compile_path),
                prep._status_paths(receipt["baseline"]["after_overlay"]["status"]),
            )

    def test_rejects_wrong_baseline_before_sparse_materialization(self) -> None:
        fixture = b"reviewed fixture\n"
        compile_path = Path("test-data/compile-fixture-tracked.bin")
        compile_value = b"tracked compile fixture\n"
        compile_fixtures = (
            (compile_path, len(compile_value), hashlib.sha256(compile_value).hexdigest()),
        )
        with tempfile.TemporaryDirectory(prefix="ppt-prep-sparse-identity-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            old_values = (
                prep.BASELINE_REF,
                prep.BASELINE_TREE,
                prep.BASELINE_LOCK_BYTES,
                prep.BASELINE_LOCK_SHA256,
                prep.FIXTURE_BYTES,
                prep.FIXTURE_SHA256,
                prep.COMPILE_FIXTURES,
            )
            prep.COMPILE_FIXTURES = compile_fixtures
            try:
                baseline_ref = repository(
                    baseline,
                    baseline=True,
                    compile_fixtures={compile_path: compile_value},
                )
                candidate_ref = repository(
                    candidate,
                    baseline=False,
                    fixture=fixture,
                    compile_fixtures={compile_path: compile_value},
                )
                baseline_tree = git(baseline, "rev-parse", "HEAD^{tree}")
                subprocess.run(
                    ["git", "-C", str(baseline), "update-index", "--skip-worktree", "--", str(compile_path)],
                    check=True,
                )
                (baseline / compile_path).unlink()
                prep.BASELINE_REF = "0" * 40
                prep.BASELINE_TREE = baseline_tree
                prep.BASELINE_LOCK_BYTES = len(b"baseline lock\n")
                prep.BASELINE_LOCK_SHA256 = hashlib.sha256(b"baseline lock\n").hexdigest()
                prep.FIXTURE_BYTES = len(fixture)
                prep.FIXTURE_SHA256 = hashlib.sha256(fixture).hexdigest()
                with self.assertRaisesRegex(prep.PreparationError, "pinned pre-lazy parent"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                        output=root / "preparation.json",
                    )
                self.assertFalse((baseline / compile_path).exists())
                self.assertEqual(git(baseline, "status", "--porcelain=v1"), "")
            finally:
                (
                    prep.BASELINE_REF,
                    prep.BASELINE_TREE,
                    prep.BASELINE_LOCK_BYTES,
                    prep.BASELINE_LOCK_SHA256,
                    prep.FIXTURE_BYTES,
                    prep.FIXTURE_SHA256,
                    prep.COMPILE_FIXTURES,
                ) = old_values

    def test_rejects_preexisting_untracked_compile_fixture(self) -> None:
        fixture = b"reviewed fixture\n"
        compile_path = Path("test-data/compile-fixture-untracked.bin")
        compile_value = b"untracked compile fixture\n"
        compile_fixtures = (
            (compile_path, len(compile_value), hashlib.sha256(compile_value).hexdigest()),
        )
        with tempfile.TemporaryDirectory(prefix="ppt-prep-untracked-") as directory:
            root = Path(directory)
            baseline = root / "baseline"
            candidate = root / "candidate"
            old_values = (
                prep.BASELINE_REF,
                prep.BASELINE_TREE,
                prep.BASELINE_LOCK_BYTES,
                prep.BASELINE_LOCK_SHA256,
                prep.FIXTURE_BYTES,
                prep.FIXTURE_SHA256,
                prep.COMPILE_FIXTURES,
            )
            prep.COMPILE_FIXTURES = compile_fixtures
            try:
                baseline_ref = repository(baseline, baseline=True)
                candidate_ref = repository(
                    candidate,
                    baseline=False,
                    fixture=fixture,
                    compile_fixtures={compile_path: compile_value},
                )
                baseline_tree = git(baseline, "rev-parse", "HEAD^{tree}")
                destination = baseline / compile_path
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(compile_value)
                prep.BASELINE_REF = baseline_ref
                prep.BASELINE_TREE = baseline_tree
                prep.BASELINE_LOCK_BYTES = len(b"baseline lock\n")
                prep.BASELINE_LOCK_SHA256 = hashlib.sha256(b"baseline lock\n").hexdigest()
                prep.FIXTURE_BYTES = len(fixture)
                prep.FIXTURE_SHA256 = hashlib.sha256(fixture).hexdigest()
                with self.assertRaisesRegex(prep.PreparationError, "outside the pinned tracked tree"):
                    prep.prepare(
                        baseline_root=baseline,
                        candidate_root=candidate,
                        candidate_ref=candidate_ref,
                        candidate_lock_sha256=hashlib.sha256(b"candidate lock\n").hexdigest(),
                        output=root / "preparation.json",
                    )
                self.assertEqual(destination.read_bytes(), compile_value)
            finally:
                (
                    prep.BASELINE_REF,
                    prep.BASELINE_TREE,
                    prep.BASELINE_LOCK_BYTES,
                    prep.BASELINE_LOCK_SHA256,
                    prep.FIXTURE_BYTES,
                    prep.FIXTURE_SHA256,
                    prep.COMPILE_FIXTURES,
                ) = old_values


if __name__ == "__main__":
    unittest.main()
