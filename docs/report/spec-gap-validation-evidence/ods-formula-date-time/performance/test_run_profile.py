"""Adversarial checks for the date/time performance runner.

These tests exercise the runner's custody and snapshot guards without invoking
Cargo or a profiling capture.  The candidate checkout deliberately has the
frozen preparation commit as HEAD: the candidate source is injected into that
checkout after the preparation commit is made.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch


RUNNER_PATH = Path(__file__).with_name("run_profile.py")
RUNNER_SPEC = importlib.util.spec_from_file_location("date_time_run_profile", RUNNER_PATH)
assert RUNNER_SPEC is not None and RUNNER_SPEC.loader is not None
RUNNER = importlib.util.module_from_spec(RUNNER_SPEC)
RUNNER_SPEC.loader.exec_module(RUNNER)


class CandidateFreezeTests(unittest.TestCase):
    """Keep freeze verification bound to the injected candidate inputs."""

    def setUp(self) -> None:
        temporary = TemporaryDirectory(prefix="litchi-date-time-runner-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.candidate = self.directory / "candidate"
        self.candidate.mkdir()

        self.baseline_path = self.directory / "baseline.json"
        self.baseline_path.write_text(
            json.dumps({"production_commit": "production", "preparation_commit": "preparation"})
            + "\n",
            encoding="utf-8",
        )
        self.patches = [
            patch.object(RUNNER, "BASELINE_PATH", self.baseline_path),
            patch.object(RUNNER, "BASELINE_COMMIT", "production"),
            patch.object(RUNNER, "PREPARATION_COMMIT", "preparation"),
        ]
        for replacement in self.patches:
            replacement.start()
            self.addCleanup(replacement.stop)

        self.selected: dict[str, str] = {}
        for relative in ("Cargo.toml", "Cargo.lock", "src/main.rs"):
            path = self.candidate / RUNNER.REL_HARNESS / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(f"fixture {relative}\n", encoding="utf-8")
            self.selected[str(RUNNER.REL_HARNESS / relative)] = RUNNER.digest(path)

    def write_freeze(self, **overrides: object) -> Path:
        freeze = {
            "source_root": str(self.candidate.resolve()),
            "selected_files": dict(self.selected),
            "base_commit": "preparation",
            "production_commit": "production",
            "candidate_commit": "injected-candidate",
        }
        freeze.update(overrides)
        path = self.directory / "freeze.json"
        path.write_text(json.dumps(freeze) + "\n", encoding="utf-8")
        return path

    def verify(self, freeze: Path, *, head: str = "preparation") -> dict[str, object]:
        with (
            patch.object(RUNNER.subprocess, "check_output", return_value="preparation\n"),
            patch.object(RUNNER, "git_head", return_value=head),
        ):
            return RUNNER.verify_candidate_freeze(self.candidate, freeze)

    def test_candidate_head_must_be_frozen_preparation_commit(self) -> None:
        freeze = self.write_freeze()

        with self.assertRaisesRegex(RuntimeError, "candidate checkout HEAD differs from frozen preparation base"):
            self.verify(freeze, head="injected-candidate")

    def test_candidate_commit_metadata_does_not_replace_preparation_head_check(self) -> None:
        """The injected candidate may be named separately from checkout HEAD."""
        freeze = self.write_freeze(candidate_commit="different-candidate-commit")

        verified = self.verify(freeze)

        self.assertEqual(verified["candidate_commit"], "different-candidate-commit")

    def test_freeze_must_bind_each_harness_input(self) -> None:
        selected = dict(self.selected)
        selected.pop(str(RUNNER.REL_HARNESS / "src/main.rs"))
        freeze = self.write_freeze(selected_files=selected)

        with self.assertRaisesRegex(RuntimeError, "candidate freeze does not bind harness input"):
            self.verify(freeze)

    def test_changed_selected_harness_input_is_rejected(self) -> None:
        freeze = self.write_freeze()
        changed = self.candidate / RUNNER.REL_HARNESS / "src/main.rs"
        changed.write_text("fixture changed after freeze\n", encoding="utf-8")

        with self.assertRaisesRegex(RuntimeError, "candidate freeze hash mismatch"):
            self.verify(freeze)


class SnapshotGuardTests(unittest.TestCase):
    """Ensure every mutable custody/hash component is checked after capture."""

    SNAPSHOT_FIELDS = (
        "git_head",
        "source_sha256",
        "workspace_source_sha256",
        "workspace_lock_sha256",
        "harness_sha256",
        "profile_input_sha256",
    )

    def test_any_snapshot_component_change_is_rejected(self) -> None:
        before = {field: f"before-{field}" for field in self.SNAPSHOT_FIELDS}

        for field in self.SNAPSHOT_FIELDS:
            with self.subTest(field=field):
                after = dict(before)
                after[field] = f"after-{field}"

                with self.assertRaisesRegex(RuntimeError, f"{field} changed during date/time capture"):
                    RUNNER.assert_unchanged(before, after, "date/time capture")


class CaptureGuardTests(unittest.TestCase):
    """Reject a mismatched harness before allocating a build target."""

    def test_mismatched_expected_harness_refuses_before_cargo_or_target(self) -> None:
        with TemporaryDirectory(prefix="litchi-date-time-capture-") as directory:
            output_dir = Path(directory) / "capture"
            snapshot = {
                "harness_sha256": {
                    "Cargo.toml": "actual-manifest",
                    "Cargo.lock": "actual-lock",
                    "src/main.rs": "actual-source",
                }
            }
            expected_harness = {
                "Cargo.toml": "frozen-manifest",
                "Cargo.lock": "frozen-lock",
                "src/main.rs": "frozen-source",
            }
            with (
                patch.object(RUNNER, "source_snapshot", return_value=snapshot),
                patch.object(RUNNER, "run_checked") as cargo,
                patch.object(RUNNER.tempfile, "mkdtemp") as target,
            ):
                with self.assertRaisesRegex(
                    RuntimeError,
                    "candidate-final harness differs from the frozen candidate harness",
                ):
                    RUNNER.capture_tree(
                        label="candidate-final",
                        source_root=Path(directory) / "candidate",
                        output_dir=output_dir,
                        warmups=0,
                        samples=1,
                        requested_cases=None,
                        profile_hashes={},
                        expected_harness=expected_harness,
                    )

            cargo.assert_not_called()
            target.assert_not_called()


if __name__ == "__main__":
    unittest.main()
