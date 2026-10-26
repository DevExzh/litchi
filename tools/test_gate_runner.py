"""Tests for the coordinator gate runner (tools/gate, record 0365)."""
from __future__ import annotations

import os
from pathlib import Path
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
RUNNER = ROOT / "tools/gate/run_gate.sh"
SHIM = ROOT / "tools/gate/bin/cargo"


def list_rows(env: dict[str, str] | None = None) -> list[str]:
    output = subprocess.check_output(
        ["/bin/bash", str(RUNNER), "--list"], cwd="/", env=env, text=True
    )
    return [line.split(" ", 1)[1] for line in output.splitlines()]


class GateRunnerTests(unittest.TestCase):
    def test_row_table_has_53_unique_rows_at_stable_indices(self) -> None:
        rows = list_rows()
        self.assertEqual(len(rows), 53)
        self.assertEqual(len(set(rows)), 53)
        expected = {
            0: "fmt",
            1: "clippy-all-features",
            3: "clippy-ops-none",
            24: "clippy-ops-numbers",
            25: "tests-ops",
            28: "clippy-leaves",
            38: "python-bindings",
            41: "fuzz",
            46: "api-glossary",
            47: "iwork-leaf-api",
            48: "tooling-unit",
            49: "crate-boundaries",
            50: "non-iwork-gate",
            51: "bench-matrix",
            52: "corpus-dispositions",
        }
        for index, name in expected.items():
            self.assertEqual(rows[index], name, index)

    def test_runner_needs_no_login_environment(self) -> None:
        # A systemd scope may start with no HOME and a minimal PATH.
        self.assertEqual(len(list_rows(env={})), 53)

    def test_every_fuzz_workspace_has_a_tracked_lock(self) -> None:
        tracked = set(
            subprocess.check_output(
                ["git", "ls-files", "crates/*/fuzz/Cargo.toml", "crates/*/fuzz/Cargo.lock"],
                cwd=ROOT,
                text=True,
            ).split()
        )
        manifests = {p for p in tracked if p.endswith("Cargo.toml")}
        self.assertTrue(manifests)
        for manifest in manifests:
            self.assertIn(manifest.replace("Cargo.toml", "Cargo.lock"), tracked)

    def test_shim_execs_the_configured_real_cargo(self) -> None:
        with tempfile.TemporaryDirectory() as scratch:
            fake = Path(scratch) / "cargo"
            fake.write_text('#!/bin/sh\necho "real:$*:${LITCHI_CAPPED-unset}"\n')
            fake.chmod(0o755)
            env = {"PATH": "/usr/bin:/bin", "LITCHI_REAL_CARGO": str(fake), "LITCHI_GATE_NO_SCOPE": "1"}
            output = subprocess.check_output([str(SHIM), "check", "--locked"], env=env, text=True)
            self.assertEqual(output.strip(), "real:check --locked:unset")
            env["LITCHI_REAL_CARGO"] = str(Path(scratch) / "missing")
            result = subprocess.run([str(SHIM), "-V"], env=env, capture_output=True, text=True)
            self.assertEqual(result.returncode, 127)
            self.assertIn("real cargo not found", result.stderr)

    def test_shim_is_first_on_the_gate_path(self) -> None:
        text = RUNNER.read_text()
        self.assertIn('export PATH="$GATE_DIR/tools/gate/bin:$CARGO_HOME/bin:', text)
        self.assertTrue(os.access(SHIM, os.X_OK))


if __name__ == "__main__":
    unittest.main()
