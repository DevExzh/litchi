#!/usr/bin/env python3
"""Failure-case tests for the source-only durable Ink qualification guards."""
import copy
import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest import mock


HERE = Path(__file__).resolve().parent


def load_script(name):
    path = HERE / name
    spec = importlib.util.spec_from_file_location(name.replace("-", "_"), path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


QUALIFY = load_script("ink-qualify.py")
ABBA = load_script("ink-abba.py")
COLLECT = load_script("ink-collect.py")


def value(route="noop", count=10, ns=500_000_000):
    return {
        "route": route,
        "iterations": count,
        "ns": ns,
        "checksum": 1,
        "output_bytes": 2,
        "output_sha256": "3" * 64,
    }


class DurableInkQualificationTests(unittest.TestCase):
    def test_snapshot_requires_a_pinned_commit(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError):
                QUALIFY.snapshot(Path(directory), "base", "HEAD", Path("/missing-lock"))

    def test_build_refs_reject_endpoint_commit_drift(self):
        build = {
            "resolved_refs": {"base": "a" * 40, "candidate": "b" * 40},
            "endpoints": {
                "base": {"source": {"commit": "a" * 40}},
                "candidate": {"source": {"commit": "b" * 40}},
            },
        }
        COLLECT.validate_build_refs(build)
        drifted = copy.deepcopy(build)
        drifted["endpoints"]["candidate"]["source"]["commit"] = "c" * 40
        with self.assertRaises(AssertionError):
            COLLECT.validate_build_refs(drifted)

    def test_calibration_must_reach_target_and_samples_minimum(self):
        case = {
            "route": "noop",
            "count": 10,
            "calibrations": [{"A": value(ns=499_999_999), "B": value()}],
            "blocks": [],
        }
        with self.assertRaises(AssertionError):
            COLLECT.validate_calibration_case(case, 500_000_000, 150_000_000)

        case["calibrations"][0]["A"]["ns"] = 500_000_000
        case["blocks"] = [{
            "legs": [
                {"side": "A", "value": value(ns=150_000_000)},
                {"side": "B", "value": value(ns=149_999_999)},
                {"side": "B", "value": value(ns=150_000_000)},
                {"side": "A", "value": value(ns=150_000_000)},
            ]
        }]
        with self.assertRaises(AssertionError):
            COLLECT.validate_calibration_case(case, 500_000_000, 150_000_000)

    def test_process_guard_fails_closed_for_unknown_and_permission_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            vanished = root / "123"
            self.assertIsNone(ABBA.process_row(vanished))

            unknown = root / "456"
            unknown.mkdir()
            with self.assertRaises(RuntimeError):
                ABBA.process_row(unknown)

            denied = root / "789"
            denied.mkdir()
            (denied / "stat").write_text("789 (probe) S 1 2\n")
            (denied / "cmdline").write_bytes(b"probe\0")
            with mock.patch.object(ABBA.os, "readlink", side_effect=PermissionError):
                with self.assertRaises(RuntimeError):
                    ABBA.process_row(denied)

    def test_preflight_has_independent_noop_replay(self):
        source = (HERE / "ink-probe.rs").read_text()
        self.assertIn("let matching_noop = apply(&matching_base, &fixture.noop);", source)
        self.assertIn("assert_eq!(encode(&matching_noop), fixture.original);", source)


if __name__ == "__main__":
    unittest.main()
