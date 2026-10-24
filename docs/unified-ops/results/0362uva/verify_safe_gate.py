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
        tooling_modules = tooling["argv"][3:]
        expected_tooling_modules = [
            # The original nine tooling-unit modules.
            "tools.test_check_api_glossary",
            "tools.test_final_api_inventory",
            "tools.test_officecli_corpus",
            "tools.test_officecli_capability_delta",
            "tools.test_no_3rdparty_references",
            "tools.test_check_crate_boundaries",
            "tools.test_non_iwork_gate",
            "tools.test_unified_bench_matrix",
            "tools.test_check_iwork_leaf_public_api",
            # Persistent inventory/disposition and W8 provenance coverage.
            "tools.test_api_debt_inventory",
            "tools.test_validate_api_disposition",
            "tools.test_w8_file_provenance",
            "tools.test_w8_original_baseline",
            "tools.test_w8_projection_policy",
        ]
        assert tooling_modules == [
            "tools.test_check_api_glossary",
            "tools.test_api_debt_inventory",
            "tools.test_validate_api_disposition",
            "tools.test_final_api_inventory",
            "tools.test_w8_file_provenance",
            "tools.test_w8_original_baseline",
            "tools.test_w8_projection_policy",
            "tools.test_officecli_corpus",
            "tools.test_officecli_capability_delta",
            "tools.test_no_3rdparty_references",
            "tools.test_check_crate_boundaries",
            "tools.test_non_iwork_gate",
            "tools.test_unified_bench_matrix",
            "tools.test_check_iwork_leaf_public_api",
        ]
        assert len(tooling_modules) == len(expected_tooling_modules) == 14
        assert len(set(tooling_modules)) == len(tooling_modules)
        assert set(expected_tooling_modules) == set(tooling_modules)
        previous = os.environ.get("LITCHI_CAPPED")
        try:
            os.environ["LITCHI_CAPPED"] = "1"
            assert "LITCHI_CAPPED" not in gate.cargo_env(args.root, target, state)
        finally:
            if previous is None:
                os.environ.pop("LITCHI_CAPPED", None)
            else:
                os.environ["LITCHI_CAPPED"] = previous

        # A /proc entry can disappear between glob and its first read. The
        # scanner skips only that ESRCH race rather than masking unrelated
        # process-table failures.
        original_path = gate.Path

        class VanishingEntry:
            name = "424242"

            def __truediv__(self, child):
                return self

            def read_bytes(self):
                raise ProcessLookupError("mock vanished PID")

        class VanishingProc:
            def glob(self, pattern):
                assert pattern == "[0-9]*"
                return [VanishingEntry()]

        def mock_path(value):
            if value == "/proc":
                return VanishingProc()
            return original_path(value)

        try:
            gate.Path = mock_path
            assert gate.process_table() == {}
        finally:
            gate.Path = original_path

        class UnexpectedEntry(VanishingEntry):
            def read_bytes(self):
                raise RuntimeError("mock unrelated process-table failure")

        class UnexpectedProc:
            def glob(self, pattern):
                return [UnexpectedEntry()]

        try:
            gate.Path = lambda value: UnexpectedProc() if value == "/proc" else original_path(value)
            try:
                gate.process_table()
            except RuntimeError as error:
                assert "unrelated process-table failure" in str(error)
            else:
                raise AssertionError("unexpected process-table failure was hidden")
        finally:
            gate.Path = original_path

        # A real child can exit while an inherited helper continues touching
        # the target. The runner must wait for a stable target sample even
        # when this is the final step; this child is wholly owned by the test
        # and exits by itself.
        original_preflight = gate.preflight
        original_source_state = gate.source_state
        source = {
            "commit": "a" * 40,
            "status": "",
            "branch": "fixture",
            "cargo_locks": {},
        }
        gate.preflight = lambda root, target, expected, isolated=False: {
            "started": 0.0,
            "target_bytes": gate.target_bytes(target),
            "free_bytes": 1 << 40,
            "builders": [],
        }
        gate.source_state = lambda root: source
        late_writer = (
            "import os,time\n"
            "target_fd=int(os.environ['LITCHI_TEST_TARGET_FD'])\n"
            "pid=os.fork()\n"
            "if pid:\n"
            "    os._exit(0)\n"
            "os.setsid()\n"
            "os.chdir('/')\n"
            "os.close(1); os.close(2)\n"
            "time.sleep(0.25)\n"
            "fd=os.open('late-child', os.O_WRONLY|os.O_CREAT|os.O_TRUNC, 0o600, dir_fd=target_fd)\n"
            "os.write(fd,b'x')\n"
            "os.close(fd)\n"
            "time.sleep(0.10)\n"
            "os._exit(0)\n"
        )
        target_fd = os.open(target, os.O_RDONLY | os.O_DIRECTORY)
        try:
            controlled = gate.direct(
                "controlled-final",
                [sys.executable, "-c", late_writer],
                env={"LITCHI_TEST_TARGET_FD": str(target_fd)},
            )
            controlled["pass_fds"] = (target_fd,)
            assert str(target) not in controlled["argv"]
            controlled_row = gate.run_step(
                controlled, temporary, target, state, source, 0, 1, "controlled-final"
            )
        finally:
            os.close(target_fd)
        assert controlled_row["status"] == "passed"
        assert controlled_row["attempt"] == "controlled-final"
        assert (target / "late-child").read_bytes() == b"x"
        sizes = [sample["target_bytes"] for sample in controlled_row["post_exit"]["samples"]]
        assert sizes[0] < sizes[-1]
        assert controlled_row["post_exit"]["target_bytes"] == gate.target_bytes(target)

        # A direct child that existed before the step is not owned merely
        # because it remains a child of the runner. A newly adopted child is
        # owned even when it has no source/target path in cwd or argv.
        runner_pid = os.getpid()
        preexisting = {
            9001: {
                "pid": 9001, "ppid": runner_pid, "pgrp": 9001,
                "start_time": 10, "command": "preexisting", "executable": "python3", "cwd": "/",
            }
        }
        launch = {
            9002: {
                "pid": 9002, "ppid": runner_pid, "pgrp": 9002,
                "start_time": 11, "command": "step", "executable": "python3", "cwd": "/",
            }
        }
        ownership = gate.new_ownership(9002, temporary, target, preexisting, launch)
        synthetic = {
            **launch,
            9001: preexisting[9001],
            9003: {
                "pid": 9003, "ppid": 9001, "pgrp": 9001,
                "start_time": 12, "command": "preexisting-child", "executable": "python3", "cwd": "/",
            },
            9004: {
                "pid": 9004, "ppid": runner_pid, "pgrp": 9004,
                "start_time": 13, "command": "detached-helper", "executable": "python3", "cwd": "/",
            },
        }
        gate._capture_owned(ownership, synthetic)
        owned_ids = {identity[0] for identity in ownership["known"]}
        assert 9004 in owned_ids
        assert 9001 not in owned_ids and 9003 not in owned_ids

        # Exercise the owned-builder branch without launching a compiler or
        # touching a foreign process. A synthetic process table is sufficient
        # to prove that the process-group identity is fail-closed.
        original_process_table = gate.process_table
        gate.process_table = lambda: {
            4242: {
                "pid": 4242,
                "ppid": 1,
                "pgrp": 31337,
                "command": "rustc --fixture",
                "executable": "rustc",
                "cwd": str(temporary),
            }
        }
        try:
            try:
                gate.post_exit_guard(31337, temporary, target)
            except RuntimeError as error:
                assert "owned builder remained" in str(error)
            else:
                raise AssertionError("owned builder was not rejected")
        finally:
            gate.process_table = original_process_table
            gate.preflight = original_preflight
            gate.source_state = original_source_state

        # A blocked retry gets a fresh current log and cannot reuse the first
        # attempt's path or bytes. The fake step never launches a child.
        retry_state = temporary / "retry-state"
        retry_target = temporary / "retry" / "target"
        retry_target.parent.mkdir()
        retry_target.mkdir()
        attempts = []
        original_make_steps = gate.make_steps
        original_run_step = gate.run_step
        original_source_state = gate.source_state
        gate.make_steps = lambda root, state_dir, target: [gate.direct("retry", ["unused"])]
        gate.source_state = lambda root: source

        def blocked_retry(step, root, target, state_dir, expected, index, total, attempt):
            log, _ = gate.attempt_paths(state_dir, index, step["name"], attempt)
            log.write_text(f"child-{attempt}\n", encoding="utf-8")
            attempts.append(attempt)
            raise RuntimeError(f"blocked-{attempt}")

        gate.run_step = blocked_retry
        previous_argv = sys.argv
        try:
            sys.argv = [
                str(args.runner), "--root", str(temporary), "--target", str(retry_target),
                "--state", str(retry_state),
            ]
            with contextlib.redirect_stdout(io.StringIO()):
                assert gate.main() == 1
            sys.argv = [
                str(args.runner), "--root", str(temporary), "--target", str(retry_target),
                "--state", str(retry_state), "--resume",
            ]
            with contextlib.redirect_stdout(io.StringIO()):
                assert gate.main() == 1
        finally:
            sys.argv = previous_argv
            gate.make_steps = original_make_steps
            gate.run_step = original_run_step
            gate.source_state = original_source_state
        assert len(attempts) == 2 and attempts[0] != attempts[1]
        retry_result = json.loads((retry_state / "state.json").read_text())
        retry_row = retry_result["rows"][0]
        assert retry_row["attempt"] == attempts[1]
        assert retry_row["log"] != str(
            gate.attempt_paths(retry_state, 0, "retry", attempts[0])[0]
        )
        assert Path(retry_row["log"]).read_text() == f"child-{attempts[1]}\n"
        assert Path(
            gate.attempt_paths(retry_state, 0, "retry", attempts[0])[0]
        ).read_text() == f"child-{attempts[0]}\n"

        # Simulate the real failure boundary: child output was already written,
        # then resource/provenance processing raised after the command ended.
        child_output = b"preserved child diagnostic\n"

        def fail_after_output(step, root, target, state_dir, source, index, total, attempt):
            log, _ = gate.attempt_paths(state_dir, index, step["name"], attempt)
            log.write_bytes(child_output)
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
        assert row["attempt"]
        assert Path(row["log"]).read_bytes() == child_output
        assert row["log_sha256"] == hashlib.sha256(child_output).hexdigest()
        failure = Path(row["failure_log"]).read_bytes()
        assert b"injected post-command resource failure" in failure
        assert row["failure_log_sha256"] == hashlib.sha256(failure).hexdigest()
        assert row["failure_log_bytes"] == len(failure)
    print("PASS: 53 steps, inventory tests, cap isolation, preserved child/failure receipts")


if __name__ == "__main__":
    main()
