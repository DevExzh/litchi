"""Exercise the public execution loop without Cargo or native probes."""
import importlib.util
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location("diagnostic", Path(__file__).with_name("run_q472a_rest29.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


class ExecuteTests(unittest.TestCase):
    def exercise(self, failure=None, barrier_failure=False, drift=False):
        with tempfile.TemporaryDirectory() as name:
            base = Path(name)
            state = base / "state"
            state.mkdir()
            compact = {"commit": m.Q_COMMIT, "branch": m.Q_BRANCH, "status": "", "cargo_locks": {}}
            rich = {"commit": m.Q_COMMIT, "locks": {}, "changed_rust_paths": []}
            seen = []
            preserved = []

            def run_step(step, root, target, phase_state, source, index, total, attempt):
                self.assertEqual(source, compact)
                self.assertIn("cargo_locks", source)
                if step["name"] == "full38":
                    self.assertTrue(preserved)
                if step["name"] == "full37":
                    source = target / "tmp/officecli-corpus-outcomes.tsv"
                    source.parent.mkdir(parents=True)
                    source.write_text("format\tcase\n")
                seen.append(step["name"])
                row = {"status": "passed", "effective_exit": 0, "name": step["name"]}
                if len(seen) == 1 and failure:
                    row.update(failure)
                return row

            def preserve(target, state):
                if barrier_failure:
                    raise RuntimeError("missing TSV")
                preserved.append(True)
                return m.load_v4().preserve_officecli(target, state)

            runner = SimpleNamespace(run_step=run_step, builders=lambda *a: [])
            v4 = SimpleNamespace(preserve_officecli=preserve)
            details = {"source": rich, "runner_source": compact, "all_steps": {
                "full53": [{"name": f"full{i}"} for i in range(53)],
                "supplements": [{"name": f"runtime{i}"} for i in range(12)],
            }}
            plan = {"state": str(state), "target": str(base / "target"), "source": rich, "skipped_prefix": {}}
            identity = m.DiagnosticError("source drift") if drift else None
            with patch.object(m, "verify_plan", return_value=(runner, v4, details)), patch.object(m, "assert_identity", side_effect=identity):
                code = m.execute(base, base / "packet", plan, "fixed")
            result = json.loads((state / "result.json").read_text())
            return code, seen, result

    def test_real_loop_passes_compact_runner_identity(self):
        code, seen, result = self.exercise()
        self.assertEqual(code, 0)
        self.assertEqual(len(seen), 29)
        self.assertEqual(result["attempted_count"], 29)
        self.assertFalse(result["acceptance"])

    def test_code_failure_continues_all_remaining_rows(self):
        code, seen, result = self.exercise({"status": "failed", "effective_exit": 101})
        self.assertEqual((code, len(seen), result["ordinary_failure_count"]), (1, 29, 1))

    def test_resource_failure_stops_before_next_step(self):
        code, seen, result = self.exercise({"status": "failed", "effective_exit": 101, "resource_failure": "target cap"})
        self.assertEqual((code, len(seen)), (2, 1))
        self.assertEqual(result["status"], "diagnostic-stopped")

    def test_barrier_failure_never_reaches_row38(self):
        code, seen, result = self.exercise(barrier_failure=True)
        self.assertEqual(code, 2)
        self.assertEqual(seen[-1], "full37")
        self.assertNotIn("full38", seen)

    def test_drift_stops_before_any_command(self):
        code, seen, result = self.exercise(drift=True)
        self.assertEqual((code, seen), (2, []))
        self.assertEqual(result["hard_stop"]["continuation_class"], "provenance")

    def test_signal_unknown_status_and_nested_holders_stop(self):
        for row in [
            {"status": "failed", "effective_exit": 137},
            {"status": "blocked", "effective_exit": 0},
            {"status": "passed", "effective_exit": 1},
            {"status": "passed", "effective_exit": 0, "post_exit": {"samples": [{"owned": [{"pid": 1}]}]}},
        ]:
            self.assertEqual(m.continuation_action(row), "stop")

    def test_cargo_oom_log_stops_even_with_exit101(self):
        with tempfile.TemporaryDirectory() as name:
            p = Path(name) / "log"
            p.write_text("compiler failed (signal: 9, SIGKILL: kill)\n")
            row = {"status": "failed", "effective_exit": 101, "log": str(p)}
            self.assertEqual(m.classify_row(row), "resource")


if __name__ == "__main__":
    unittest.main()
