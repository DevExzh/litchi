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
VALIDATOR = load(
    "ppt_correctness_validator_under_test",
    ROOT / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py",
)


class PptGateWiringTests(unittest.TestCase):
    def test_ppt_steps_use_host_guard_and_shared_release_target(self) -> None:
        with tempfile.TemporaryDirectory(prefix="ppt-gate-steps-") as temporary:
            root = Path(temporary)
            (root / "crates/litchi-ops").mkdir(parents=True)
            (root / "crates/litchi-ops/Cargo.toml").write_text(
                "[features]\nppt = []\n", encoding="utf-8"
            )
            state = root / "state"
            target = root / "target"
            steps = {step["name"]: step for step in GATE.make_steps(root, state, target)}

        build = steps["ppt-scoped-build"]
        suite = steps["ppt-scoped-correctness-suite"]
        correctness = steps["ppt-scoped-correctness"]
        receipt = steps["ppt-scoped-correctness-validate"]
        self.assertTrue(build["isolated"])
        self.assertTrue(suite["isolated"])
        self.assertTrue(correctness["isolated"])
        self.assertTrue(receipt["isolated"])
        self.assertEqual(build["argv"][:3], ["python3", str(root / "docs/unified-ops/results/0362g/ppt_host_process_guard.py"), "--"])
        self.assertIn("scoped_patch_probe", build["argv"])
        self.assertIn("scoped_patch_correctness", build["argv"])
        self.assertEqual(suite["argv"][:3], build["argv"][:3])
        suite_command = suite["argv"][-1]
        self.assertEqual(suite["env"], {"RUST_TEST_THREADS": "1"})
        self.assertEqual(suite_command.count("--test-threads=1"), 5)
        for package in ("litchi-ppt", "litchi-cfb"):
            self.assertIn(f"-p {package}", suite_command)
        self.assertIn("cargo clippy --offline --locked -p litchi-ppt", suite_command)
        self.assertIn("cargo clippy --offline --locked -p litchi-cfb", suite_command)
        self.assertEqual(correctness["argv"][:3], build["argv"][:3])
        self.assertEqual(
            correctness["argv"][3],
            str(root / "target/release/examples/scoped_patch_correctness"),
        )
        self.assertEqual(receipt["argv"][1], str(root / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"))

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
            row["external_media_path"] = "other.avi"
            receipt.write_text(json.dumps(row), encoding="utf-8")
            with self.assertRaises(SystemExit):
                VALIDATOR.main(["validator", str(receipt), str(fixture)])


if __name__ == "__main__":
    unittest.main()
