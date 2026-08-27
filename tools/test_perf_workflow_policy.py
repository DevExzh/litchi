"""Static resource-safety policy checks for the performance workflow."""

from __future__ import annotations

import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "perf-baseline.yml"


def _yaml_scalar(value: str) -> str:
    """Return a comparable scalar without YAML quotes or comments."""
    value = value.strip()
    if " #" in value:
        value = value.split(" #", 1)[0].rstrip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        return value[1:-1]
    return value


def _env_assignments(text: str, key: str) -> list[tuple[int, str]]:
    """Read key/value lines that are children of actual YAML ``env`` maps."""
    pattern = re.compile(rf"^(?P<indent> *){re.escape(key)}\s*:\s*(?P<value>.*)$")
    assignments: list[tuple[int, str]] = []
    lines = text.splitlines()
    env_pattern = re.compile(r"^(?P<indent> *)env:\s*(?:#.*)?$")
    for index, line in enumerate(lines):
        env_match = env_pattern.match(line)
        if env_match is None:
            continue
        env_indent = len(env_match.group("indent"))
        child_indent: int | None = None
        for child_index in range(index + 1, len(lines)):
            following = lines[child_index]
            if not following.strip():
                continue
            following_indent = len(following) - len(following.lstrip(" "))
            if following_indent <= env_indent:
                break
            if child_indent is None:
                child_indent = following_indent
            if following_indent != child_indent:
                continue
            match = pattern.match(following)
            if match is None:
                continue
            value = _yaml_scalar(match.group("value"))
            if value in {"|", "|-", "|+", ">", ">-", ">+"}:
                continuation: list[str] = []
                for continuation_line in lines[child_index + 1 :]:
                    if continuation_line.strip() and (
                        len(continuation_line)
                        - len(continuation_line.lstrip(" "))
                        <= child_indent
                    ):
                        break
                    if continuation_line.strip():
                        continuation.append(continuation_line.strip())
                value = " ".join(continuation)
            assignments.append((following_indent, value))
    return assignments


def _job_blocks(text: str) -> dict[str, str]:
    lines = text.splitlines()
    jobs_index = next(
        index
        for index, line in enumerate(lines)
        if re.fullmatch(r"jobs:\s*", line)
    )
    header = re.compile(r"^  (?P<name>[A-Za-z0-9][A-Za-z0-9_-]*):\s*(?:#.*)?$")
    starts: list[tuple[int, str]] = []
    for index in range(jobs_index + 1, len(lines)):
        line = lines[index]
        if line and not line.startswith(" "):
            break
        match = header.match(line)
        if match is not None:
            starts.append((index, match.group("name")))
    blocks: dict[str, str] = {}
    for position, (start, name) in enumerate(starts):
        end = starts[position + 1][0] if position + 1 < len(starts) else len(lines)
        blocks[name] = "\n".join(lines[start:end])
    return blocks


def _global_env_values(text: str, key: str) -> list[str]:
    jobs_index = next(
        index
        for index, line in enumerate(text.splitlines())
        if re.fullmatch(r"jobs:\s*", line)
    )
    return [value for _, value in _env_assignments("\n".join(text.splitlines()[:jobs_index]), key)]


def _steps(job: str) -> list[str]:
    lines = job.splitlines()
    steps_index = next(
        (index for index, line in enumerate(lines) if re.fullmatch(r"\s*steps:\s*", line)),
        None,
    )
    if steps_index is None:
        return []
    steps_indent = len(lines[steps_index]) - len(lines[steps_index].lstrip(" ")) + 2
    starts = [
        index
        for index in range(steps_index + 1, len(lines))
        if len(lines[index]) - len(lines[index].lstrip(" ")) == steps_indent
        and lines[index].lstrip().startswith("-")
    ]
    return [
        "\n".join(lines[start : starts[position + 1] if position + 1 < len(starts) else len(lines)])
        for position, start in enumerate(starts)
    ]


def _run_body(step: str) -> str:
    lines = step.splitlines()
    if not lines:
        return ""
    step_indent = len(lines[0]) - len(lines[0].lstrip(" "))
    field_indent = step_indent + 2
    run_pattern = re.compile(rf"^ {{{field_indent}}}run:\s*(?P<inline>.*)$")
    for index, line in enumerate(lines):
        match = run_pattern.match(line)
        if match is None:
            continue
        body = [match.group("inline")]
        for continuation in lines[index + 1 :]:
            if continuation.strip() and (
                len(continuation) - len(continuation.lstrip(" ")) <= field_indent
            ):
                break
            body.append(continuation)
        return "\n".join(body)
    return ""


def _has_always_condition(step: str) -> bool:
    lines = step.splitlines()
    if not lines:
        return False
    step_indent = len(lines[0]) - len(lines[0].lstrip(" "))
    field_pattern = re.compile(
        rf"^ {{{step_indent + 2}}}if:\s*(?P<value>.*?)(?:\s+#.*)?$"
    )
    return any(
        (match := field_pattern.match(line)) is not None
        and re.search(r"\balways\s*\(\s*\)", match.group("value")) is not None
        for line in lines[1:]
    )


def _top_level_section(text: str, name: str) -> str:
    lines = text.splitlines()
    header = f"  {name}:"
    start = next(index for index, line in enumerate(lines) if line == header)
    end = len(lines)
    for index in range(start + 1, len(lines)):
        if re.fullmatch(r"  [A-Za-z0-9_-]+:\s*", lines[index]):
            end = index
            break
    return "\n".join(lines[start:end])


def _has_forbidden_recursive_rm(text: str) -> bool:
    flattened = re.sub(r"\\\s*\n", " ", text)
    for command in re.split(r"[;&|\n]", flattened):
        command = command.split("#", 1)[0]
        match = re.search(r"\brm\b(?P<arguments>.*)$", command)
        if match is None:
            continue
        options = re.findall(
            r"(?<![A-Za-z0-9_])(--[A-Za-z][A-Za-z0-9-]*|-[A-Za-z]+)",
            match.group("arguments"),
        )
        recursive = any(
            option == "--recursive"
            or (option.startswith("-") and not option.startswith("--") and "r" in option[1:].lower())
            for option in options
        )
        force = any(
            option == "--force"
            or (option.startswith("-") and not option.startswith("--") and "f" in option[1:].lower())
            for option in options
        )
        if recursive and force:
            return True
    return False


def _explicit_cargo_job_counts(text: str) -> list[str]:
    pattern = re.compile(
        r"(?<![\w-])--jobs(?:\s*=\s*|\s+)([0-9]+)\b"
        r"|(?<![\w-])-j(?:\s*=\s*|\s*)([0-9]+)\b"
    )
    return [first or second for first, second in pattern.findall(text)]


def _has_cargo_command(job: str) -> bool:
    return any(
        re.search(r"(?:^|[;&|])\s*cargo(?=\s|$)", _run_body(step), re.MULTILINE)
        is not None
        for step in _steps(job)
    )


_CARGO_TARGET_VARIABLE = r"(?:\$CARGO_TARGET_DIR|\$\{CARGO_TARGET_DIR\})"
_CARGO_TARGET_OPERAND = rf"(?:\"{_CARGO_TARGET_VARIABLE}\"|'${{CARGO_TARGET_DIR}}'|{_CARGO_TARGET_VARIABLE})"


def _has_exact_target_guard(step: str, target_value: str) -> bool:
    target = re.escape(target_value)
    variable = rf"(?:\"?{_CARGO_TARGET_VARIABLE}\"?)"
    return re.search(
        rf"(?:\[\s*|\btest\s+){variable}\s*=\s*\"?{target}\"?\s*\]?",
        step,
    ) is not None


def _shell_target_value(target_value: str) -> str:
    return re.sub(
        r"\$\{\{\s*runner\.temp\s*\}\}",
        "$RUNNER_TEMP",
        target_value,
    )


def _has_non_symlink_guard(step: str) -> bool:
    return re.search(
        rf"!\s*-L\s+\"?{_CARGO_TARGET_VARIABLE}\"?",
        step,
    ) is not None


CARGO_WORKLOAD_JOBS = ("smoke", "full")


class PerformanceWorkflowPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.workflow = WORKFLOW_PATH.read_text(encoding="utf-8")
        cls.jobs = _job_blocks(cls.workflow)

    def test_workflow_has_serial_cargo_builds_and_disabled_incremental_artifacts(self) -> None:
        build_values = [
            value for _, value in _env_assignments(self.workflow, "CARGO_BUILD_JOBS")
        ]
        incremental_values = [
            value for _, value in _env_assignments(self.workflow, "CARGO_INCREMENTAL")
        ]
        self.assertTrue(build_values, "workflow must set CARGO_BUILD_JOBS explicitly")
        self.assertTrue(
            incremental_values,
            "workflow must set CARGO_INCREMENTAL explicitly",
        )
        self.assertTrue(all(value.strip("\"'") == "1" for value in build_values))
        self.assertTrue(all(value.strip("\"'") == "0" for value in incremental_values))

        explicit_job_counts = _explicit_cargo_job_counts(self.workflow)
        self.assertTrue(all(count == "1" for count in explicit_job_counts))

    def test_cargo_target_dir_is_explicitly_runner_temp_backed(self) -> None:
        target_values = [
            value for _, value in _env_assignments(self.workflow, "CARGO_TARGET_DIR")
        ]
        self.assertTrue(
            target_values,
            "workflow must set CARGO_TARGET_DIR explicitly",
        )
        for value in target_values:
            self.assertRegex(
                value,
                r"\$\{\{\s*runner\.temp\s*\}\}",
                msg=f"CARGO_TARGET_DIR is not runner-temp backed: {value!r}",
            )
        self.assertEqual(len(set(target_values)), 1)

    def test_cargo_workload_jobs_are_explicit_and_complete(self) -> None:
        self.assertEqual(set(CARGO_WORKLOAD_JOBS), {"smoke", "full"})
        for name in CARGO_WORKLOAD_JOBS:
            self.assertIn(name, self.jobs)
        detected = {
            name
            for name, job in self.jobs.items()
            if _has_cargo_command(job)
        }
        self.assertEqual(
            detected,
            set(CARGO_WORKLOAD_JOBS),
            "a Cargo-bearing job must be added to CARGO_WORKLOAD_JOBS before it can bypass policy",
        )

    def test_every_job_has_a_positive_timeout(self) -> None:
        self.assertTrue(self.jobs, "workflow must declare at least one job")
        for name, job in self.jobs.items():
            timeout_values = [
                int(match.group(1))
                for match in re.finditer(
                    r"^    timeout-minutes:\s*[\"']?([0-9]+)[\"']?\s*(?:#.*)?$",
                    job,
                    re.MULTILINE,
                )
            ]
            self.assertEqual(
                len(timeout_values),
                1,
                f"performance job {name!r} must have one job-level timeout-minutes",
            )
            self.assertGreater(timeout_values[0], 0)

    def test_every_job_has_always_on_resource_diagnostics(self) -> None:
        for name in CARGO_WORKLOAD_JOBS:
            job = self.jobs[name]
            always_steps = "\n".join(
                step for step in _steps(job) if _has_always_condition(step)
            )
            self.assertRegex(
                always_steps,
                r"\bdf\s+(?:-[^\n\s]*h\b|--human-readable\b)",
                msg=f"performance job {name!r} lacks always-on disk diagnostics",
            )
            self.assertRegex(
                always_steps,
                r"(?:\bfree\s+(?:-[^\n\s]*h\b|--human-readable\b)|/proc/meminfo\b)",
                msg=f"performance job {name!r} lacks always-on memory diagnostics",
            )
            self.assertRegex(
                always_steps,
                r"\bdu\s+(?:-[^\n\s]*s[^\n\s]*\b|--summarize\b)",
                msg=f"performance job {name!r} lacks always-on target-size diagnostics",
            )

    def test_cargo_workload_jobs_clean_exact_runner_temp_without_crossing_mounts(self) -> None:
        self.assertFalse(_has_forbidden_recursive_rm(self.workflow))
        find_pattern = re.compile(
            rf"\bfind\s+{_CARGO_TARGET_OPERAND}\s+-xdev\b[^\n]*-depth[^\n]*-delete\b"
        )
        global_target_values = _global_env_values(self.workflow, "CARGO_TARGET_DIR")
        self.assertLessEqual(len(global_target_values), 1)
        for name in CARGO_WORKLOAD_JOBS:
            job = self.jobs[name]
            local_target_values = [
                value for _, value in _env_assignments(job, "CARGO_TARGET_DIR")
            ]
            resolved_target_values = local_target_values or global_target_values
            self.assertEqual(
                len(resolved_target_values),
                1,
                f"performance job {name!r} must resolve one Cargo target directory",
            )
            target_value = resolved_target_values[0]
            self.assertRegex(target_value, r"\$\{\{\s*runner\.temp\s*\}\}")
            shell_target_value = _shell_target_value(target_value)
            cleanup_steps = [step for step in _steps(job) if find_pattern.search(step)]
            self.assertTrue(
                cleanup_steps,
                f"performance job {name!r} lacks find -xdev -depth -delete cleanup",
            )
            self.assertTrue(
                any(_has_always_condition(step) for step in cleanup_steps),
                f"performance job {name!r} cleanup must run with always()",
            )
            self.assertTrue(
                any(
                    _has_exact_target_guard(step, shell_target_value)
                    for step in cleanup_steps
                ),
                f"performance job {name!r} cleanup must compare the exact runner-temp target",
            )
            self.assertTrue(
                any(_has_non_symlink_guard(step) for step in cleanup_steps),
                f"performance job {name!r} cleanup must reject symlink targets",
            )

    def test_artifact_uploads_are_failure_visible(self) -> None:
        uploads = [
            step
            for job in self.jobs.values()
            for step in _steps(job)
            if re.search(r"uses:\s*actions/upload-artifact@", step)
        ]
        self.assertTrue(uploads, "workflow must define its artifact uploads in job steps")
        for step in uploads:
            self.assertTrue(
                _has_always_condition(step),
                "every performance artifact upload must use if: always()",
            )

    def test_stale_repository_target_cache_is_not_configured(self) -> None:
        self.assertNotRegex(
            self.workflow,
            r"tools/perf-baseline[\\/]target(?:[\\/]|\b)",
            "the retired repository-local perf target must not be cached or used",
        )

    def test_policy_file_is_triggered_and_run_by_smoke(self) -> None:
        for event in ("push", "pull_request"):
            section = _top_level_section(self.workflow, event)
            self.assertRegex(
                section,
                r"(?m)^\s*-\s*['\"]?tools/test_perf_workflow_policy\.py['\"]?\s*(?:#.*)?$",
                msg=f"{event} path filter must include the workflow policy test",
            )
        smoke_steps = _steps(self.jobs["smoke"])
        self.assertTrue(
            any(
                re.search(r"python3\s+-m\s+unittest", step)
                and "tools.test_perf_workflow_policy" in step
                for step in smoke_steps
            ),
            "smoke must execute the workflow policy test",
        )


if __name__ == "__main__":
    unittest.main()
