import importlib.util
from pathlib import Path
import tempfile
import unittest

spec = importlib.util.spec_from_file_location("diagnostic", Path(__file__).with_name("run_q472a_rest29.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


class OfficeCliBarrierTests(unittest.TestCase):
    def exercise(self, kind):
        with tempfile.TemporaryDirectory() as name:
            base = Path(name)
            target, state = base / "target", base / "state"
            target.mkdir()
            state.mkdir()
            outside = base / "outside"
            outside.mkdir()
            source = target / "tmp/officecli-corpus-outcomes.tsv"
            if kind == "source-parent-link":
                (target / "tmp").symlink_to(outside, target_is_directory=True)
            else:
                source.parent.mkdir()
            if kind == "source-link":
                external = outside / "report.tsv"
                external.write_text("format\tcase\n")
                source.symlink_to(external)
            else:
                source.write_text("format\tcase\n")
            source.chmod(0o640)
            if kind == "destination-parent-link":
                (state / "officecli-row37").symlink_to(outside, target_is_directory=True)
            row = {"status": "passed", "effective_exit": 0}
            if kind == "regular":
                receipt = m.preserve_officecli(m.load_v4(), target, state, row)
                self.assertEqual(receipt["mode"], "0640")
                self.assertEqual(Path(receipt["destination"]).read_bytes(), source.read_bytes())
            else:
                with self.assertRaises(m.DiagnosticError):
                    m.preserve_officecli(m.load_v4(), target, state, row)
                self.assertFalse((outside / "preservation.json").exists())

    def test_regular_output_preserves_bytes_and_mode(self):
        self.exercise("regular")

    def test_source_symlink_refused(self):
        self.exercise("source-link")

    def test_source_parent_symlink_refused(self):
        self.exercise("source-parent-link")

    def test_destination_parent_symlink_refused(self):
        self.exercise("destination-parent-link")


if __name__ == "__main__":
    unittest.main()
