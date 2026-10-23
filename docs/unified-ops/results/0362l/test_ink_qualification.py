#!/usr/bin/env python3
"""Failure-case tests for the source-only durable Ink qualification guards."""
import copy
import importlib.util
import json
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


def write_collector_fixture(root):
    """Build a small complete collector manifest with a nonempty source delta."""
    root.mkdir()
    zero = root / "zero"
    zero.write_bytes(b"")
    probe = root / "probe.rs"
    probe.write_text("fn main() {}\n")
    lock = root / "Cargo.lock"
    lock.write_text("# fixed test lock\n")
    binary = root / "ink-native"
    binary.write_bytes(b"protected probe\n")
    project_manifest = root / "Cargo.toml"
    project_manifest.write_text("[profile.release]\nlto = true\npanic = \"abort\"\n")

    shared = {
        "crates/litchi-core/same.rs": b"same\n",
        "crates/litchi-opc/changed.rs": b"before\n",
        "crates/litchi-ooxml-common/removed.rs": b"removed\n",
    }
    candidate_shared = {
        "crates/litchi-core/same.rs": b"same\n",
        "crates/litchi-opc/changed.rs": b"after\n",
        "crates/litchi-ooxml-common/added.rs": b"added\n",
    }
    commits = {"base": "a" * 40, "candidate": "b" * 40}
    source_rows = {}
    source_manifests = {}
    for label, files in (("base", shared), ("candidate", candidate_shared)):
        source = root / f"source-{label}"
        for relative, contents in files.items():
            path = source / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(contents)
        archive = root / f"source-{label}.tar"
        archive.write_bytes(f"archive-{label}\n".encode())
        rows = []
        for path in sorted(source.rglob("*")):
            if path.is_file():
                row = COLLECT.artifact(path)
                row["path"] = str(path.relative_to(source))
                rows.append(row)
        manifest_path = root / f"source-{label}.json"
        manifest_path.write_text(json.dumps({
            "commit": commits[label],
            "archive": COLLECT.artifact(archive),
            "files": rows,
        }, indent=2) + "\n")
        source_rows[label] = {row["path"]: (row["bytes"], row["sha256"]) for row in rows}
        source_manifests[label] = manifest_path

    validation = root / "validation"
    validation.mkdir()
    stream = {
        "schema": COLLECT.STREAM_SCHEMA,
        "read_chunk": 4096,
        "sink_window": 8192,
        "base_bytes": 10,
        "target_bytes": 11,
        "routes": [
            {"route": route, "bytes": 10, "sha256": "a" * 64}
            for route in ("forward", "inverse", "noop")
        ],
    }
    output_bytes = {
        "forward.docx": b"forward",
        "forward.json": b"forward-json",
        "inverse.docx": b"inverse",
        "inverse.json": b"inverse-json",
        "noop.docx": b"noop",
        "noop.json": b"noop-json",
        "source.docx": b"source",
        "streaming.json": (json.dumps(stream) + "\n").encode(),
    }
    for name, contents in output_bytes.items():
        (validation / name).write_bytes(contents)
    validation_rows = [COLLECT.artifact(validation / name) for name in sorted(output_bytes)]

    endpoints = {}
    for label in ("base", "candidate"):
        command = {"exit_code": 0, "stdout": COLLECT.artifact(zero),
                   "stderr": COLLECT.artifact(zero)}
        endpoints[label] = {
            "source": {"commit": commits[label],
                        "manifest": COLLECT.artifact(source_manifests[label])},
            "project_manifest": COLLECT.artifact(project_manifest),
            "project_source": COLLECT.artifact(probe),
            "lock": COLLECT.artifact(lock),
            "registry": [],
            "rustc": "rustc test\n",
            "cargo": "cargo test\n",
            "binaries": {"native": COLLECT.artifact(binary),
                         "allocation": COLLECT.artifact(binary)},
            "commands": [command],
            "validation": {"native": validation_rows, "allocation": validation_rows},
        }
    delta = QUALIFY.source_delta(
        source_rows["base"], source_rows["candidate"], QUALIFY_SHARED,
        commits["base"], commits["candidate"])
    build = {
        "schema": COLLECT.QUALIFY_SCHEMA,
        "workload": "docx-durable-ink",
        "comparison": COLLECT.COMPARISON,
        "status": "complete",
        "builder": COLLECT.artifact(zero),
        "probe": COLLECT.artifact(probe),
        "input_lock": COLLECT.artifact(lock),
        "resolved_refs": commits,
        "endpoints": endpoints,
        "shared_source_delta": delta,
    }
    (root / "build.json").write_text(json.dumps(build, indent=2) + "\n")


QUALIFY_SHARED = (
    "crates/litchi-core/",
    "crates/litchi-opc/",
    "crates/litchi-ooxml-common/",
    "crates/litchi-drawingml/",
)


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

    def test_shared_source_delta_records_changes_and_missing_files(self):
        base = {
            "crates/litchi-opc/old.rs": (3, "a"),
            "crates/litchi-opc/changed.rs": (4, "b"),
        }
        candidate = {
            "crates/litchi-opc/changed.rs": (5, "c"),
            "crates/litchi-opc/new.rs": (6, "d"),
        }
        delta = QUALIFY.source_delta(base, candidate, ("crates/litchi-opc/",), "a" * 40, "b" * 40)
        self.assertEqual(delta["base_commit"], "a" * 40)
        self.assertEqual(delta["candidate_commit"], "b" * 40)
        self.assertEqual(delta["base_only"], [{"path": "crates/litchi-opc/old.rs",
                                               "base": {"bytes": 3, "sha256": "a"}}])
        self.assertEqual(delta["candidate_only"], [{"path": "crates/litchi-opc/new.rs",
                                                    "candidate": {"bytes": 6, "sha256": "d"}}])
        self.assertEqual(delta["changed"][0], {
            "path": "crates/litchi-opc/changed.rs",
            "base": {"bytes": 4, "sha256": "b"},
            "candidate": {"bytes": 5, "sha256": "c"},
        })

    def test_collector_accepts_and_checks_nonempty_shared_delta(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "qualification"
            write_collector_fixture(root)
            collected = COLLECT.collect_build(root)
            self.assertEqual(collected["comparison"], "integrated-route")
            delta = collected["shared_source_delta"]
            self.assertEqual(len(delta["changed"]), 1)
            self.assertEqual(len(delta["base_only"]), 1)
            self.assertEqual(len(delta["candidate_only"]), 1)

            build_path = root / "build.json"
            build = json.loads(build_path.read_text())
            build["shared_source_delta"]["changed"][0]["base"]["sha256"] = "0" * 64
            build_path.write_text(json.dumps(build, indent=2) + "\n")
            with self.assertRaises(AssertionError):
                COLLECT.collect_build(root)


if __name__ == "__main__":
    unittest.main()
