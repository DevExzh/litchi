"""Synthetic receipt tests; these are not benchmark evidence."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).parent
SPEC = importlib.util.spec_from_file_location("collector", ROOT / "constructor-collect.py")
COLLECTOR = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(COLLECTOR)


class ResumeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.binaries = {}
        for side in "AB":
            path = self.directory / side
            path.write_text("synthetic test artifact\n")
            self.binaries[side] = COLLECTOR.artifact(path)
        self.build = {"endpoints": {
            label: {"binaries": {"native": self.binaries[side]}}
            for side, label in (("A", "base"), ("B", "candidate"))}}

    def collect(self):
        return COLLECTOR.collect_native(self.directory, self.build)[0]

    def save(self, path, value):
        path.write_text(json.dumps(value))

    def attempt(self, index, blocks, completed_before=0, blocked=False):
        directory = self.directory / f"native-a{index}"
        directory.mkdir()
        (self.directory / f"native-a{index}.log").write_text("synthetic test\n")
        requested = [{"route": route, "blocks": 5 - completed_before if n == 0 else 5}
                     for n, route in enumerate(COLLECTOR.ROUTES)
                     if n or completed_before < 5]
        case = {"route": COLLECTOR.ROUTES[0], "count": 1, "blocks": []}
        for block in range(blocks):
            legs = []
            for position, side in enumerate("ABBA"):
                label = f"leg-{block}-{position}-{side}"
                value = dict(route=case["route"], iterations=1, ns=500_000_000,
                             checksum=1, output_bytes=1, output_sha256="test")
                stdout, stderr = directory / f"{label}.stdout", directory / f"{label}.stderr"
                self.save(stdout, value)
                stderr.write_text("")
                self.save(directory / f"{label}.receipt.json", dict(
                    command=["taskset", "-c", "29", self.binaries[side]["path"], case["route"], "1"],
                    exit_code=0, overlap=None, stdout=COLLECTOR.artifact(stdout),
                    stderr=COLLECTOR.artifact(stderr)))
                for suffix in ("before", "during-0001", "after"):
                    self.save(directory / f"{label}-{suffix}.host.json", dict(
                        processes=[{"synthetic": True}], blockers=["overlap"] if blocked else []))
                legs.append(dict(side=side, label=label, value=value))
            case["blocks"].append(dict(legs=legs, ratio=1.0))
        manifest = dict(schema="constructor-abba.v2", binaries=self.binaries,
                        runner=COLLECTOR.artifact(ROOT / "constructor-abba.py"),
                        planner=COLLECTOR.artifact(ROOT / "constructor-collect.py"),
                        cpu=29, target_ns=500_000_000, minimum_ns=150_000_000,
                        requested=requested, cases=[case], status="running")
        self.save(directory / "manifest.json", manifest)
        return directory, manifest

    def test_empty_is_not_qualified(self):
        result = self.collect()
        self.assertFalse(result["qualified"])
        self.assertTrue(all(not row["blocks"] for row in result["cases"]))

    def test_resume_keeps_every_completed_block(self):
        self.attempt(1, 3)
        self.attempt(2, 2, completed_before=3)
        result = self.collect()
        self.assertEqual(len(result["cases"][0]["blocks"]), 5)
        self.assertEqual(result["cases"][0]["median_ratio"], 1.0)
        self.assertFalse(result["qualified"])

    def test_restarting_completed_blocks_is_rejected(self):
        self.attempt(1, 3)
        self.attempt(2, 2)
        with self.assertRaises(AssertionError):
            self.collect()

    def test_skipping_a_route_is_rejected(self):
        directory, manifest = self.attempt(1, 0)
        manifest["requested"] = manifest["requested"][1:]
        self.save(directory / "manifest.json", manifest)
        with self.assertRaises(AssertionError):
            self.collect()

    def test_missing_attempt_is_rejected(self):
        self.attempt(2, 0)
        with self.assertRaises(AssertionError):
            self.collect()

    def test_overlap_cannot_qualify_a_block(self):
        self.attempt(1, 1, blocked=True)
        with self.assertRaises(AssertionError):
            self.collect()

    def test_tampered_output_is_rejected(self):
        directory, _ = self.attempt(1, 1)
        next(directory.glob("*.stdout")).write_text("tampered")
        with self.assertRaises(AssertionError):
            self.collect()

    def test_incomplete_attempt_cannot_claim_completion(self):
        directory, manifest = self.attempt(1, 1)
        manifest["status"] = "complete"
        self.save(directory / "manifest.json", manifest)
        with self.assertRaises(AssertionError):
            self.collect()


if __name__ == "__main__":
    unittest.main()
