from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]


def load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GATE = load("ppt_run_safe_gate_under_test", HERE / "run_safe_gate.py")
LOCK_GUARD = load(
    "ppt_candidate_lock_guard_under_test",
    ROOT / "docs/unified-ops/results/0362g/validate_ppt_candidate_lock.py",
)
VALIDATOR = load(
    "ppt_correctness_validator_under_test",
    ROOT / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py",
)


EXPECTED_NORMAL_STEP_NAMES = (
    "fmt",
    "clippy-all-features",
    "clippy-ops-no-encryption",
    "clippy-ops-",
    "clippy-ops-xlsx",
    "clippy-ops-pptx",
    "clippy-ops-docx",
    "clippy-ops-xls",
    "clippy-ops-xlsb",
    "clippy-ops-doc",
    "clippy-ops-ppt",
    "clippy-ops-rtf",
    "clippy-ops-odt",
    "clippy-ops-ods",
    "clippy-ops-odp",
    "clippy-ops-oth",
    "clippy-ops-odm",
    "clippy-ops-odg",
    "clippy-ops-odc",
    "clippy-ops-odi",
    "clippy-ops-odb",
    "clippy-ops-odf-formula",
    "clippy-ops-pages",
    "clippy-ops-keynote",
    "clippy-ops-numbers",
    "tests-ops",
    "tests-facade",
    "tests-core",
    "clippy-leaves",
    "clippy-ods",
    "tests-leaves",
    "tests-ods-full",
    "clippy-workspace",
    "tests-workspace-excluding-ods-ops",
    "wsdoc",
    "ppt-alone",
    "rustdoc-core-ops-facade",
    "cli-tests",
    "python-bindings",
    "feature-matrix",
    "feature-matrix-clippy",
    "fuzz",
    "goldens",
    "golden-clean",
    "perf-harness-check",
    "native-resave-check",
    "api-glossary",
    "iwork-leaf-api",
    "tooling-unit",
    "crate-boundaries",
    "non-iwork-gate",
    "bench-matrix",
    "corpus-dispositions",
)


class PptGateWiringTests(unittest.TestCase):
    def test_lock_hashes_includes_dedicated_odf_consumer_lock(self) -> None:
        with tempfile.TemporaryDirectory(prefix="odf-consumer-lock-hash-") as temporary:
            root = Path(temporary)
            lock = root / "tools/odf-raw-codec-consumer/Cargo.lock"
            lock.parent.mkdir(parents=True)
            lock.write_bytes(b"dedicated odf consumer lock")
            self.assertEqual(
                GATE.lock_hashes(root),
                {"tools/odf-raw-codec-consumer/Cargo.lock": hashlib.sha256(lock.read_bytes()).hexdigest()},
            )

    def test_ppt_steps_use_host_guard_and_shared_release_target(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-gate-steps-") as temporary:
            state = Path(temporary) / "state"
            target = Path(temporary) / "target"
            normal = GATE.plan_steps(ROOT, state, target, set())
            auxiliary = GATE.plan_steps(
                ROOT,
                state,
                target,
                set(GATE.PPT_AUXILIARY_NAMES),
            )
            self.assertEqual(len(normal), 53)
            self.assertEqual(
                tuple(step["name"] for step in normal),
                EXPECTED_NORMAL_STEP_NAMES,
            )
            self.assertEqual(len(auxiliary), 57)
            self.assertFalse(any(step["name"].startswith("ppt-scoped-") for step in normal))
            steps = {step["name"]: step for step in auxiliary}

        build = steps["ppt-scoped-build"]
        suite = steps["ppt-scoped-correctness-suite"]
        correctness = steps["ppt-scoped-correctness"]
        receipt = steps["ppt-scoped-correctness-validate"]
        self.assertTrue(build["isolated"])
        self.assertTrue(suite["isolated"])
        self.assertTrue(correctness["isolated"])
        self.assertTrue(receipt["isolated"])
        lock_guard = ROOT / "docs/unified-ops/results/0362g/validate_ppt_candidate_lock.py"
        host_guard = ROOT / "docs/unified-ops/results/0362g/ppt_host_process_guard.py"
        self.assertEqual(build["argv"][:2], ["python3", str(lock_guard)])
        self.assertIn(str(host_guard), build["argv"])
        self.assertIn("scoped_patch_probe", build["argv"])
        self.assertIn("scoped_patch_correctness", build["argv"])
        self.assertEqual(suite["argv"][:2], build["argv"][:2])
        suite_command = suite["argv"][-1]
        self.assertEqual(suite["env"], {"RUST_TEST_THREADS": "1"})
        self.assertEqual(suite_command.count("--test-threads=1"), 5)
        for package in ("litchi-ppt", "litchi-cfb"):
            self.assertIn(f"-p {package}", suite_command)
        self.assertIn("cargo clippy --offline --locked -p litchi-ppt", suite_command)
        self.assertIn("cargo clippy --offline --locked -p litchi-cfb", suite_command)
        self.assertEqual(correctness["argv"][:3], ["python3", str(host_guard), "--"])
        self.assertEqual(
            correctness["argv"][3],
            str(target / "release/examples/scoped_patch_correctness"),
        )
        self.assertEqual(receipt["argv"][1], str(ROOT / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"))

    def test_candidate_lock_guard_rejects_wrong_bytes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-lock-guard-") as temporary:
            lock = Path(temporary) / "Cargo.lock"
            lock.write_bytes(b"not the approved candidate lock")
            with self.assertRaises(SystemExit):
                LOCK_GUARD.main(
                    [
                        "validate_ppt_candidate_lock.py",
                        str(lock),
                        GATE.PPT_CANDIDATE_LOCK_SHA256,
                    ]
                )
            with self.assertRaises(SystemExit):
                LOCK_GUARD.main(
                    [
                        "validate_ppt_candidate_lock.py",
                        str(lock),
                        "0" * 64,
                    ]
                )

    def test_receipt_validator_accepts_changed_target_and_inverse(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-correctness-receipt-") as temporary:
            root = Path(temporary)
            fixture = root / "video.ppt"
            fixture.write_bytes(b"source fixture")
            source_hash = hashlib.sha256(fixture.read_bytes()).hexdigest()
            receipt = root / "receipt.json"
            receipt.write_text(
                json.dumps(
                    {
                        "mode": "scoped_patch_correctness",
                        "fixture": str(fixture),
                        "media_id": 1,
                        "requested_external_media_path": "clip-0000.avi",
                        "external_media_path": "clip-0000.avi",
                        "source_len": fixture.stat().st_size,
                        "target_len": fixture.stat().st_size + 1,
                        "source_sha256": source_hash,
                        "target_sha256": "b" * 64,
                        "replay_target_sha256": "b" * 64,
                        "replay_matches_target": True,
                        "inverse_source_sha256": source_hash,
                        "inverse_matches_source": True,
                    }
                ),
                encoding="utf-8",
            )
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                self.assertEqual(VALIDATOR.main(["validator", str(receipt), str(fixture)]), 0)
            self.assertIn('"status": "passed"', output.getvalue())

            row = json.loads(receipt.read_text(encoding="utf-8"))
            row["target_sha256"] = source_hash
            receipt.write_text(json.dumps(row), encoding="utf-8")
            with self.assertRaises(SystemExit):
                VALIDATOR.main(["validator", str(receipt), str(fixture)])

            row["target_sha256"] = "b" * 64
            row["replay_target_sha256"] = "c" * 64
            row["replay_matches_target"] = True
            receipt.write_text(json.dumps(row), encoding="utf-8")
            with self.assertRaises(SystemExit):
                VALIDATOR.main(["validator", str(receipt), str(fixture)])

            row["replay_target_sha256"] = "b" * 64
            row["replay_matches_target"] = False
            receipt.write_text(json.dumps(row), encoding="utf-8")
            with self.assertRaises(SystemExit):
                VALIDATOR.main(["validator", str(receipt), str(fixture)])

            row["target_sha256"] = "b" * 64
            row["external_media_path"] = "other.avi"
            receipt.write_text(json.dumps(row), encoding="utf-8")
            with self.assertRaises(SystemExit):
                VALIDATOR.main(["validator", str(receipt), str(fixture)])


if __name__ == "__main__":
    unittest.main()
