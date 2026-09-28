"""Root-owned source mirror tests for the 0806 constructor-reuse amendment.

The script is deliberately executable only by the root coordinator.  It
constructs two minimal workspaces from the packet's six-file source archives,
then runs the same 70/100 test and warning-denied Clippy gates used by 0805.
No production file is copied into either workspace at runtime.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
OWNERS = tuple(PLAN["quality"]["helper_crates"])
EXPECTED_TESTS = {"before": PLAN["quality"]["before_test_count"], "after": PLAN["quality"]["after_test_count"]}
TEST_RESULT = re.compile(r"test result: ok\. (\d+) passed; (\d+) failed;")


def source_digest_map(leg: str) -> dict[str, str]:
    return {
        str(path.relative_to(P / "source" / leg)): c.sha(path)
        for path in sorted((P / "source" / leg).iterdir())
        if path.is_file()
    }


def write_project(leg: str) -> Path:
    project = P / "test-src" / leg
    if project.exists():
        raise AssertionError(f"quality project already exists: {project}")
    project.mkdir(parents=True)
    members = ",".join(f'"crates/{owner}"' for owner in OWNERS)
    (project / "Cargo.toml").write_text(
        f'[workspace]\nresolver="3"\nmembers=[{members}]\n', encoding="utf-8"
    )
    source_dir = P / "source" / leg
    for owner in OWNERS:
        crate = project / "crates" / owner
        src = crate / "src"
        src.mkdir(parents=True)
        (crate / "Cargo.toml").write_text(
            "[package]\n"
            f'name = "{owner}"\n'
            'version = "0.0.0"\n'
            'edition = "2024"\n'
            'publish = false\n'
            "[dependencies]\n"
            'quick-xml = "=0.41.0"\n',
            encoding="utf-8",
        )
        (src / "lib.rs").write_text(
            "#![forbid(unsafe_code)]\n#![allow(dead_code)]\nmod xml_attributes;\n",
            encoding="utf-8",
        )
        module = source_dir / f"{owner}-xml_attributes.rs"
        shutil.copyfile(module, src / "xml_attributes.rs")
    tests = project / "crates/litchi-opc/src/xml_attributes/tests.rs"
    tests.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source_dir / "litchi-opc-xml_attributes-tests.rs", tests)
    return project


def tests_passed(log: Path) -> tuple[int, int]:
    passed = failed = 0
    for line in log.read_text(encoding="utf-8", errors="replace").splitlines():
        match = TEST_RESULT.search(line)
        if match:
            passed += int(match.group(1))
            failed += int(match.group(2))
    return passed, failed


def main() -> None:
    out = P / "quality"
    if out.exists():
        raise AssertionError(f"quality output already exists: {out}")
    out.mkdir()
    if not (P / "source.json").exists():
        c.write(P / "source.json", c.source())
    initial_archives = {leg: source_digest_map(leg) for leg in ("before", "after")}
    c.write(out / "inputs.json", {
        "source": {leg: c.relative_artifact(P / "source.json") for leg in ("before", "after")},
        "archives": initial_archives,
        "plan": c.sha(P / "plan.json"),
        "driver": c.sha(Path(__file__)),
    })
    env = os.environ | {
        "CARGO_TARGET_DIR": str(c.TARGET / "helper-tests"),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
    }
    rows: list[dict[str, object]] = []
    counts: dict[str, int] = {}
    for leg in ("before", "after"):
        project = write_project(leg)
        manifest = project / "Cargo.toml"
        commands = [
            ["cargo", "generate-lockfile", "--offline", "--manifest-path", str(manifest)],
            [
                "cargo", "test", "--offline", "--locked", "--manifest-path", str(manifest),
                "--workspace", "--", "--test-threads=2",
            ],
            [
                "cargo", "clippy", "--offline", "--locked", "--manifest-path", str(manifest),
                "--workspace", "--all-targets", "--", "-D", "warnings",
            ],
        ]
        for index, command in enumerate(commands):
            log = out / f"{leg}-{index}.log"
            started = time.time()
            with log.open("w", encoding="utf-8") as stream:
                result = subprocess.run(
                    command, cwd=c.ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT
                )
            row = {
                "leg": leg,
                "command": command,
                "started": started,
                "ended": time.time(),
                "exit_code": result.returncode,
                "log": c.relative_artifact(log),
            }
            rows.append(row)
            c.write(out / "receipts.json", rows)
            if result.returncode != 0:
                raise SystemExit(f"quality command failed: {log}")
            if index == 1:
                passed, failed = tests_passed(log)
                counts[leg] = passed
                if failed != 0 or passed != EXPECTED_TESTS[leg]:
                    raise AssertionError(
                        f"{leg}: expected {EXPECTED_TESTS[leg]} tests, got {passed} passed/{failed} failed"
                    )
        if initial_archives[leg] != source_digest_map(leg):
            raise AssertionError(f"{leg}: source archive changed during quality run")
    c.write(out / "complete.json", {
        "schema": "litchi.performance.0806.amendment-quality.v1",
        "rows": rows,
        "test_counts": counts,
        "expected_test_counts": EXPECTED_TESTS,
        "source_archives": initial_archives,
        "scope": "five helper mirror crates plus shared OPC tests; no full production-crate claim",
    })


if __name__ == "__main__":
    main()
