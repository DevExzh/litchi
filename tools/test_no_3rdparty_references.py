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

# These are policy predicates and negative-fixture inputs, not references to
# the absent upstream checkout.  Keep the exceptions line-scoped and exact:
# a file-wide allow-list would let an actual fixture path hide in the same
# file.  The placeholder keeps this guard from matching its own source.
_UPSTREAM_TOKEN = "3rd" + "party"


def _exact_policy_line(value: str) -> re.Pattern[str]:
    return re.compile(re.escape(value.replace("__UPSTREAM__", _UPSTREAM_TOKEN)))


INTENTIONAL_POLICY_LINES: dict[str, tuple[re.Pattern[str], ...]] = {
    "tools/native_qualification/prepare_signed_native_inputs_v2.py": (
        _exact_policy_line(
            '    _expect("__UPSTREAM__" not in path.parts, f"{label} may not use an upstream checkout")'
        ),
        _exact_policy_line(
            '    _expect("__UPSTREAM__" not in absolute.parts, f"{label} may not use an upstream checkout")'
        ),
        _exact_policy_line(
            '    _expect("__UPSTREAM__" not in archive_path.parts, '
            '"bound source archive may not use an upstream checkout")'
        ),
    ),
    "tools/raw_api_coverage.py": (
        _exact_policy_line('    if "__UPSTREAM__" in path_obj.parts:'),
        _exact_policy_line('        raise CoverageError(f"{label} must not reference __UPSTREAM__/")'),
    ),
    "tools/test_signed_native_policy_v2.py": (
        _exact_policy_line(
            '        for value in ("/tmp/input", "../input", "a/../input", "test-data/__UPSTREAM__/file"):'
        ),
    ),
}


def tracked_files() -> list[str]:
    result = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-z", "--", *SCOPES],
        check=True,
        capture_output=True,
    )
    return [name for name in result.stdout.decode().split("\0") if name]


def _is_intentional_policy_line(name: str, text: str) -> bool:
    return any(pattern.fullmatch(text) for pattern in INTENTIONAL_POLICY_LINES.get(name, ()))


def offenders() -> list[str]:
    """Matching lines of tracked files (working-tree content), minus exact guards."""
    result = subprocess.run(
        ["git", "-C", str(ROOT), "grep", "-n", "-I", "-P", NEEDLE.pattern, "--", *SCOPES],
        capture_output=True,
    )
    if result.returncode not in (0, 1):
        raise RuntimeError(result.stderr.decode(errors="replace"))
    found = []
    for line in result.stdout.decode(errors="replace").splitlines():
        name, separator, rest = line.partition(":")
        _line_number, separator, text = rest.partition(":")
        if not separator or not _is_intentional_policy_line(name, text):
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

    def test_intentional_policy_lines_are_current_and_exact(self) -> None:
        tracked = set(tracked_files())
        for name, patterns in INTENTIONAL_POLICY_LINES.items():
            with self.subTest(name=name):
                self.assertIn(name, tracked, "drop the stale policy-line entry")
                lines = (ROOT / name).read_text(encoding="utf-8", errors="replace").splitlines()
                for pattern in patterns:
                    self.assertEqual(
                        sum(pattern.fullmatch(line) is not None for line in lines),
                        1,
                        f"policy exception is stale or duplicated: {name} / {pattern.pattern}",
                    )

    def test_policy_line_exception_does_not_allow_a_fixture_path(self) -> None:
        allowed = (
            f'    _expect("{_UPSTREAM_TOKEN}" not in path.parts, '
            'f"{label} may not use an upstream checkout")'
        )
        self.assertTrue(_is_intentional_policy_line(
            "tools/native_qualification/prepare_signed_native_inputs_v2.py", allowed
        ))
        self.assertFalse(_is_intentional_policy_line(
            "tools/native_qualification/prepare_signed_native_inputs_v2.py",
            f'    source = "test-data/{_UPSTREAM_TOKEN}/file"',
        ))
        self.assertFalse(_is_intentional_policy_line(
            "tools/raw_api_coverage.py",
            f'    return root / "{_UPSTREAM_TOKEN}" / "fixture"',
        ))

    def test_the_pattern_matches_path_forms(self) -> None:
        name = "3rd" + "party"
        for line in (f'join("../../{name}/poi")', f'/ "{name}" / "x"', f"`{name}/specs`"):
            self.assertRegex(line, NEEDLE)
        self.assertNotRegex("third-party notices", NEEDLE)


if __name__ == "__main__":
    unittest.main()
