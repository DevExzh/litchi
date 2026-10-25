"""Synthetic tests for the current-source native policy v2 boundary."""

from __future__ import annotations

import copy
import hashlib
import inspect
import json
from pathlib import Path
import subprocess
import tempfile
import unittest

from native_qualification import current_native_policy_v2 as policy
from native_qualification import prepare_current_native_inputs_v2 as preparer
from native_qualification import validate_current_native_v2 as validator


ROOT = Path(__file__).resolve().parents[1]


class CurrentNativePolicyV2Tests(unittest.TestCase):
    def test_candidate_is_pending_and_execution_fails_closed(self) -> None:
        self.assertEqual(policy.TRUSTED_POLICY.status, "incomplete")
        self.assertEqual(len(policy.SOURCE_TRANSITIONS), 3)
        with self.assertRaises(policy.PolicyIncomplete):
            policy.TRUSTED_POLICY.require_execution_ready()
        with self.assertRaises(policy.PolicyIncomplete):
            policy.make_ppt_steps_v2(Path("/tmp/source"), Path("/tmp/state"), Path("/tmp/target"))

    def test_hints_are_not_accepted_as_hashes(self) -> None:
        for hint in policy.CANDIDATE_HINTS.values():
            self.assertIsNone(policy.SHA256.fullmatch(hint))
        self.assertNotIn("expected_sha256", inspect.signature(policy.validate_lock).parameters)
        self.assertNotIn("candidate_lock_sha256", inspect.signature(policy.make_ppt_steps_v2).parameters)

    def test_manifests_match_the_fixed_route_contract(self) -> None:
        for name, lane in (("ppt-current-v2.json", "ppt-auxiliary"), ("legacy-current-v2.json", "legacy-current")):
            path = ROOT / "tools/native_qualification/manifests" / name
            value = json.loads(path.read_text(encoding="utf-8"))
            policy.validate_routes(lane, value["routes"])
            self.assertEqual(value["policy_schema"], policy.POLICY_SCHEMA)

        value = json.loads((ROOT / "tools/native_qualification/manifests/ppt-current-v2.json").read_text())
        tampered = copy.deepcopy(value)
        tampered["routes"][0]["output"] = "target/release/other"
        with self.assertRaises(policy.PolicyMismatch):
            policy.validate_routes("ppt-auxiliary", tampered["routes"])

    def test_pending_preparation_is_diagnostic_only(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        with tempfile.TemporaryDirectory() as directory:
            parent = Path(directory)
            before = self._fake_source(parent / "before", "before")
            after = self._fake_source(parent / "after", "after")
            output = parent / "receipt.json"
            receipt = preparer.prepare(
                lane="ppt-auxiliary",
                manifest=manifest,
                before=before,
                after=after,
                output=output,
            )
            self.assertEqual(receipt["status"], "blocked")
            self.assertFalse(receipt["execution_admitted"])
            self.assertEqual(receipt["policy_sha256"], policy.TRUSTED_POLICY.digest())
            self.assertEqual(json.loads(output.read_text()), receipt)
            with self.assertRaises(policy.PolicyIncomplete):
                validator.validate_preparation(output)

    def test_preparation_rejects_caller_route_overlay(self) -> None:
        manifest = ROOT / "tools/native_qualification/manifests/ppt-current-v2.json"
        value = json.loads(manifest.read_text(encoding="utf-8"))
        value["routes"][1]["program"] = "caller-selected"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "manifest.json"
            path.write_bytes((json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode())
            before = self._fake_source(Path(directory) / "before", "before")
            after = self._fake_source(Path(directory) / "after", "after")
            with self.assertRaises(preparer.PreparationError):
                preparer.prepare(
                    lane="ppt-auxiliary",
                    manifest=path,
                    before=before,
                    after=after,
                    output=Path(directory) / "receipt.json",
                )

    @staticmethod
    def _fake_source(path: Path, marker: str) -> Path:
        path.mkdir(parents=True)
        (path / "Cargo.toml").write_text(f"[package]\nname = \"{marker}\"\nversion = \"0.0.0\"\n")
        subprocess.run(["git", "init", "-q"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.email", "policy-test@example.invalid"], cwd=path, check=True)
        subprocess.run(["git", "config", "user.name", "Policy Test"], cwd=path, check=True)
        subprocess.run(["git", "add", "Cargo.toml"], cwd=path, check=True)
        subprocess.run(["git", "commit", "-q", "-m", "fixture"], cwd=path, check=True)
        return path


if __name__ == "__main__":
    unittest.main()
