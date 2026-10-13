#!/usr/bin/env python3
"""Guard: no tracked code under crates/, tools/ or fuzz/ names the untracked
upstream checkout directory (AGENTS.md, section "Test data: never reference
..."; record 0257).

That directory is absent from worktrees, CI and other checkouts, so a test,
tool or comment that points into it either breaks or misleads. Fixtures are
vendored under ``test-data/`` with their provenance instead, and citations
name the upstream project and its own path.
"""

from __future__ import annotations

import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCOPES = ("crates", "tools", "fuzz")
# Built by concatenation so this guard does not match itself.
NEEDLE = re.compile(r"\b3rd" + r"party\b")

# Temporary exceptions: files that parallel branches own and are fixing.
# Each entry names its owner; delete the entry when that branch merges.
ALLOWLIST: dict[str, str] = {
}


def tracked_files() -> list[str]:
    result = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-z", "--", *SCOPES],
        check=True,
        capture_output=True,
    )
    return [name for name in result.stdout.decode().split("\0") if name]


def offenders() -> list[str]:
    """Matching lines of tracked files (working-tree content), outside the allowlist."""
    result = subprocess.run(
        ["git", "-C", str(ROOT), "grep", "-n", "-I", "-P", NEEDLE.pattern, "--", *SCOPES],
        capture_output=True,
    )
    if result.returncode not in (0, 1):
        raise RuntimeError(result.stderr.decode(errors="replace"))
    this = Path(__file__).resolve().relative_to(ROOT).as_posix()
    found = []
    for line in result.stdout.decode(errors="replace").splitlines():
        name = line.split(":", 1)[0]
        if name != this and name not in ALLOWLIST:
            found.append(line)
    return found


class NoUpstreamCheckoutReferences(unittest.TestCase):
    def test_tracked_code_does_not_reference_the_upstream_checkout(self) -> None:
        found = offenders()
        self.assertEqual(
            found,
            [],
            "vendor the fixture under test-data/ with its provenance, or cite the "
            "upstream project and its own path:\n" + "\n".join(found),
        )

    def test_allowlisted_files_still_exist_and_still_need_the_exception(self) -> None:
        tracked = set(tracked_files())
        for name in ALLOWLIST:
            with self.subTest(name=name):
                self.assertIn(name, tracked, "drop the stale allowlist entry")
                text = (ROOT / name).read_text(encoding="utf-8", errors="replace")
                self.assertRegex(text, NEEDLE, "drop the allowlist entry; the file is clean")

    def test_the_pattern_matches_path_forms(self) -> None:
        name = "3rd" + "party"
        for line in (f'join("../../{name}/poi")', f'/ "{name}" / "x"', f"`{name}/specs`"):
            self.assertRegex(line, NEEDLE)
        self.assertNotRegex("third-party notices", NEEDLE)


if __name__ == "__main__":
    unittest.main()
