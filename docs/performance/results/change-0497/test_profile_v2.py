#!/usr/bin/env python3
"""Focused parser and custody tests for the 0497 v2 syscall profile helper."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("profile_0497_v2", ROOT / "profile_v2.py")
assert SPEC is not None and SPEC.loader is not None
PROFILE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PROFILE)


class ProfileV2ConfigTests(unittest.TestCase):
    def test_filter_uses_supported_stat_syscall_set(self) -> None:
        self.assertNotIn("fstatat", PROFILE.SYSCALLS)
        self.assertIn("newfstatat", PROFILE.SYSCALLS)

    def test_cli_requires_a_subcommand(self) -> None:
        with self.assertRaises(SystemExit):
            PROFILE._parser().parse_args([])

    def test_cli_rejects_unknown_subcommand(self) -> None:
        with self.assertRaises(SystemExit):
            PROFILE._parser().parse_args(["unsupported", "attempt"])


def _atomic_trace_fixture(
    root: Path,
    *,
    reversed_order: bool = False,
    sync_before_write: bool = False,
    destination_count: int = 1,
) -> dict[str, Path]:
    if destination_count <= 0:
        raise ValueError("destination_count must be positive")
    tmpdir = root / "tmp"
    replay = root / "replay"
    report = root / "report.json"
    tmpdir.mkdir()
    replay.mkdir()
    replay_file = replay / "source8192-authored256-window-near.replay"
    publication_lines: list[str] = []
    destinations: list[Path] = []
    parents: list[Path] = []
    siblings: list[Path] = []
    for serial in range(destination_count):
        destination_parent = tmpdir / f"litchi-docx-replay-42-{serial}"
        destination_parent.mkdir()
        destination = destination_parent / "published.docx"
        sibling = destination_parent / f".litchi-abcd{serial}.tmp"
        destinations.append(destination)
        parents.append(destination_parent)
        siblings.append(sibling)
        publication_lines.extend([
            f'openat(AT_FDCWD, "{destination_parent}", O_RDONLY) = {4 + serial * 2}<{destination_parent}>',
            f'openat({4 + serial * 2}<{destination_parent}>, "{sibling.name}", O_RDWR) = {5 + serial * 2}<{sibling}>',
        ])
        if sync_before_write:
            publication_lines.extend([
                f'fsync({5 + serial * 2}<{sibling}>) = 0',
                f'write({5 + serial * 2}<{sibling}>, "x", 1) = 1',
                f'rename("{sibling}", "{destination}") = 0',
            ])
        elif reversed_order:
            publication_lines.extend([
                f'write({5 + serial * 2}<{sibling}>, "x", 1) = 1',
                f'rename("{sibling}", "{destination}") = 0',
                f'fsync({5 + serial * 2}<{sibling}>) = 0',
            ])
        else:
            publication_lines.extend([
                f'write({5 + serial * 2}<{sibling}>, "x", 1) = 1',
                f'fsync({5 + serial * 2}<{sibling}>) = 0',
                f'rename("{sibling}", "{destination}") = 0',
            ])
        publication_lines.append(f'fsync({4 + serial * 2}<{destination_parent}>) = 0')
    trace = root / "trace"
    trace.write_text(
        "\n".join([
            f'openat(AT_FDCWD, "{replay}", O_RDONLY) = 3<{replay}>',
            f'write(3<{replay_file}>, "x", 1) = 1',
            f'fdatasync(3<{replay_file}>) = 0',
            *publication_lines,
            f'write(6<{report}>, "x", 1) = 1',
        ])
        + "\n",
        encoding="utf-8",
    )
    return {
        "tmpdir": tmpdir,
        "replay": replay,
        "report": report,
        "trace": trace,
        "destination": destinations[0],
        "parent": parents[0],
        "sibling": siblings[0],
        "destinations": destinations,
        "parents": parents,
        "siblings": siblings,
    }


class ProfileTests(unittest.TestCase):
    def test_atomic_trace_preserves_paths_and_separates_sync_scopes(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root)
            tmpdir = fixture["tmpdir"]
            replay = fixture["replay"]
            report = fixture["report"]
            trace = fixture["trace"]
            destination = fixture["destination"]
            sibling = fixture["sibling"]
            parsed = PROFILE._parse_trace(
                trace, mode="atomic", tmpdir=tmpdir, replay_dir=replay, report=report,
                expected_atomic_paths=[{
                    "destination": str(destination),
                    "private_parent": str(destination.parent),
                }],
            )
            self.assertEqual(parsed["paths"]["destination"], [str(destination)])
            self.assertEqual(parsed["paths"]["sibling_temporary"], [str(sibling)])
            self.assertGreater(parsed["phases"]["authored_file_store"]["sync"]["calls"], 0)
            self.assertGreater(parsed["phases"]["output_atomic"]["sync"]["calls"], 0)
            self.assertEqual(parsed["phases"]["report"]["write"]["calls"], 1)
            self.assertEqual(parsed["atomic_lifecycle"]["event_order"][0]["ordered"], True)

    def test_atomic_trace_binds_warmup_and_report_destinations_in_order(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root, destination_count=2)
            parsed = PROFILE._parse_trace(
                fixture["trace"],
                mode="atomic",
                tmpdir=fixture["tmpdir"],
                replay_dir=fixture["replay"],
                report=fixture["report"],
                expected_atomic_paths=[{
                    "destination": str(fixture["destinations"][1]),
                    "private_parent": str(fixture["parents"][1]),
                }],
                warmups=1,
                samples=1,
            )
            self.assertEqual(
                parsed["paths"]["warmup_trace_only_destinations"],
                [str(fixture["destinations"][0])],
            )
            self.assertEqual(
                parsed["paths"]["reported_destinations"],
                [str(fixture["destinations"][1])],
            )
            self.assertEqual(len(parsed["atomic_lifecycle"]["event_order"]), 2)

    def test_atomic_trace_rejects_unexpected_extra_destination(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root, destination_count=3)
            with self.assertRaisesRegex(PROFILE.ProfileError, "destination count"):
                PROFILE._parse_trace(
                    fixture["trace"],
                    mode="atomic",
                    tmpdir=fixture["tmpdir"],
                    replay_dir=fixture["replay"],
                    report=fixture["report"],
                    expected_atomic_paths=[{
                        "destination": str(fixture["destinations"][2]),
                        "private_parent": str(fixture["parents"][2]),
                    }],
                    warmups=1,
                    samples=1,
                )

    def test_atomic_trace_rejects_report_path_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root)
            common = {
                "tmpdir": fixture["tmpdir"],
                "replay_dir": fixture["replay"],
                "report": fixture["report"],
                "mode": "atomic",
            }
            bad_bindings = (
                {
                    "destination": str(fixture["destination"].with_name("wrong.docx")),
                    "private_parent": str(fixture["parent"]),
                },
                {
                    "destination": str(fixture["destination"]),
                    "private_parent": str(root / "wrong-private-parent"),
                },
            )
            for binding in bad_bindings:
                with self.subTest(binding=binding):
                    with self.assertRaises(PROFILE.ProfileError):
                        PROFILE._parse_trace(
                            fixture["trace"],
                            expected_atomic_paths=[binding],
                            **common,
                        )

    def test_atomic_trace_rejects_reversed_publication_order(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root, reversed_order=True)
            with self.assertRaisesRegex(PROFILE.ProfileError, "sibling write/fsync order"):
                PROFILE._parse_trace(
                    fixture["trace"],
                    mode="atomic",
                    tmpdir=fixture["tmpdir"],
                    replay_dir=fixture["replay"],
                    report=fixture["report"],
                    expected_atomic_paths=[{
                        "destination": str(fixture["destination"]),
                        "private_parent": str(fixture["parent"]),
                    }],
                )

    def test_atomic_trace_rejects_fsync_before_final_sibling_write(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            fixture = _atomic_trace_fixture(root, sync_before_write=True)
            with self.assertRaisesRegex(PROFILE.ProfileError, "sibling write/fsync order"):
                PROFILE._parse_trace(
                    fixture["trace"],
                    mode="atomic",
                    tmpdir=fixture["tmpdir"],
                    replay_dir=fixture["replay"],
                    report=fixture["report"],
                    expected_atomic_paths=[{
                        "destination": str(fixture["destination"]),
                        "private_parent": str(fixture["parent"]),
                    }],
                )

    def test_counting_trace_rejects_atomic_paths_by_contract(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            tmpdir = root / "tmp"
            replay = root / "replay"
            report = root / "report.json"
            tmpdir.mkdir()
            replay.mkdir()
            replay_file = replay / "authored.replay"
            trace = root / "trace"
            trace.write_text(
                "\n".join(
                    [
                        f'openat(AT_FDCWD, "{replay_file}", O_RDWR) = 3<{replay_file}>',
                        f'write(3<{replay_file}>, "x", 1) = 1',
                        f'fdatasync(3<{replay_file}>) = 0',
                        f'write(4<{report}>, "x", 1) = 1',
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            parsed = PROFILE._parse_trace(
                trace, mode="counting", tmpdir=tmpdir, replay_dir=replay, report=report
            )
            self.assertEqual(parsed["paths"]["destination"], [])
            self.assertEqual(parsed["paths"]["sibling_temporary"], [])
            self.assertEqual(parsed["phases"]["output_atomic"]["write"]["calls"], 0)

    def test_private_cleanup_removes_only_empty_roots(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "run"
            tmpdir = root / "tmp"
            replay = root / "replay"
            tmpdir.mkdir(parents=True)
            replay.mkdir()
            receipt = PROFILE._cleanup_private(root, tmpdir, replay)
            self.assertEqual(receipt["status"], "pass")
            self.assertEqual(receipt["remaining"], [])
            self.assertFalse(root.exists())


if __name__ == "__main__":
    unittest.main()
