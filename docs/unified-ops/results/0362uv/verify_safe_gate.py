"""Build-free checks for the external coordinator gate runner repair."""
from __future__ import annotations

import argparse
import contextlib
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import sys
import tempfile


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runner", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--scratch", type=Path, required=True)
    args = parser.parse_args()
    spec = importlib.util.spec_from_file_location("safe_gate_checked", args.runner)
    assert spec is not None and spec.loader is not None
    gate = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gate)
    args.scratch.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gate-receipt-", dir=args.scratch) as directory:
        temporary = Path(directory)
        target = temporary / "target"
        target.mkdir()
        state = temporary / "state"
        steps = gate.make_steps(args.root, state, target)
        assert len(steps) == 53
        tooling = next(step for step in steps if step["name"] == "tooling-unit")
        assert "tools.test_final_api_inventory" in tooling["argv"]
        previous = os.environ.get("LITCHI_CAPPED")
        try:
            os.environ["LITCHI_CAPPED"] = "1"
            assert "LITCHI_CAPPED" not in gate.cargo_env(args.root, target, state)
        finally:
            if previous is None:
                os.environ.pop("LITCHI_CAPPED", None)
            else:
                os.environ["LITCHI_CAPPED"] = previous

        # Simulate the real failure boundary: child output was already written,
        # then resource/provenance processing raised after the command ended.
        child_output = b"preserved child diagnostic\n"

        def fail_after_output(step, root, target, state_dir, source, index, total):
            (state_dir / f"{index:02d}-{step['name']}.log").write_bytes(child_output)
            raise RuntimeError("injected post-command resource failure")

        gate.source_state = lambda root: {
            "commit": "a" * 40, "status": "", "branch": "fixture", "cargo_locks": {}
        }
        gate.make_steps = lambda root, state, target: [gate.direct("fixture", ["unused"])]
        gate.run_step = fail_after_output
        previous_argv = sys.argv
        try:
            sys.argv = [str(args.runner), "--root", str(temporary), "--target", str(target),
                        "--state", str(state)]
            with contextlib.redirect_stdout(io.StringIO()):
                assert gate.main() == 1
        finally:
            sys.argv = previous_argv
        result = json.loads((state / "state.json").read_text())
        row = result["rows"][0]
        assert row["status"] == "blocked"
        assert Path(row["log"]).read_bytes() == child_output
        assert row["log_sha256"] == hashlib.sha256(child_output).hexdigest()
        failure = Path(row["failure_log"]).read_bytes()
        assert b"injected post-command resource failure" in failure
        assert row["failure_log_sha256"] == hashlib.sha256(failure).hexdigest()
        assert row["failure_log_bytes"] == len(failure)
    print("PASS: 53 steps, inventory tests, cap isolation, preserved child/failure receipts")


if __name__ == "__main__":
    main()
