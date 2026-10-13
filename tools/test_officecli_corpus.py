from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools import officecli_corpus as corpus


def registry(elements: list[dict]) -> corpus.Registry:
    return corpus.Registry("docx", {"elements": elements, "gaps": []}, [])


PARAGRAPH = {
    "element": "paragraph",
    "operations": {"add": True, "set": True, "get": True},
    "properties": {
        "align": {"type": "enum", "add": True, "set": True, "get": True},
        "style": {"type": "string", "add": False, "set": False, "get": True},
    },
}
CHART = {
    "element": "chart",
    "operations": {},
    "properties": {},
    "gap": {"code": "unsupported_capability", "followUp": "R4x", "reason": "no reader"},
}


def call(**fields) -> corpus.Call:
    values = dict(script="word/x.sh", line=1, verb="add", fmt="docx", kind="paragraph",
                  result="ok", detail="", target="")
    values.update(fields)
    return corpus.Call(**values)


class ParseTests(unittest.TestCase):
    def test_rows_and_dashes(self):
        text = "# header\n@ word/x.sh\n3\tadd\tdocx\tparagraph\tok\t-\t-\n"
        (row,) = corpus.parse(text)
        self.assertEqual((row.script, row.line, row.kind, row.detail), ("word/x.sh", 3, "paragraph", ""))

    def test_malformed_row(self):
        with self.assertRaises(SystemExit):
            corpus.parse("@ a.sh\n1\tadd\n")

    def test_family_folds_templates(self):
        self.assertEqual(corpus.family("ppt/templates/styles/a/build.sh"), "ppt/templates")
        self.assertEqual(corpus.family("excel/charts/x.sh"), "excel/charts")


class ClassifyTests(unittest.TestCase):
    def classify(self, calls, elements=(PARAGRAPH, CHART)):
        return corpus.classify(calls, {"docx": registry(list(elements))})

    def test_pass_resident_and_translated(self):
        verdicts = self.classify([
            call(),
            call(verb="open", result="skip"),
            call(fmt="pptx", kind="slide", detail="translated:D4"),
        ])
        self.assertEqual([v.kind for v in verdicts], ["passed", "delta", "delta"])
        self.assertEqual([v.rule for v in verdicts[1:]], ["D5", "D4"])

    def test_registry_gap_agrees(self):
        (verdict,) = self.classify([call(kind="chart", result="unsupported_element")])
        self.assertEqual(verdict.kind, "gap")
        self.assertIn("R4x", verdict.note)

    def test_refused_claimed_property_is_a_bug(self):
        (verdict,) = self.classify([call(result="unsupported_property", detail="align")])
        self.assertEqual(verdict.kind, "bug")

    def test_unclaimed_property_is_a_gap(self):
        (verdict,) = self.classify([call(result="unsupported_property", detail="style")])
        self.assertEqual(verdict.kind, "gap")
        self.assertIn("not add-able", verdict.note)

    def test_cascade_needs_an_earlier_refused_add(self):
        alone = self.classify([call(verb="set", result="path_not_found", target="p")])
        self.assertEqual(alone[0].kind, "bug")
        after = self.classify([
            call(kind="chart", result="unsupported_element"),
            call(verb="set", result="path_not_found", target="p"),
        ])
        self.assertEqual(after[1].kind, "cascade")

    def test_missing_verb_and_stray_word(self):
        verdicts = self.classify([
            call(verb="validate", result="usage", detail="validate"),
            call(result="usage", detail="A,Product"),
        ])
        self.assertEqual([v.rule for v in verdicts], ["G1", "D7"])

    def test_warned_calls_and_their_skipped_properties(self):
        calls = [call(result="warned", detail="style,colour,+2")]
        regs = {"docx": registry([PARAGRAPH, CHART])}
        (verdict,) = corpus.classify(calls, regs)
        self.assertEqual((verdict.kind, verdict.rule), ("warned", "F3"))
        # A warned call is not a refusal, so it starts no cascade.
        self.assertFalse(calls[0].refused)
        skipped = corpus.skipped_verdicts(calls, regs)
        self.assertEqual([v.rule for _, v in skipped], [
            "docx|add|paragraph|warned|style",
            "docx|add|paragraph|warned|colour",
            "docx|add|paragraph|warned|(more)",
        ])
        # Skipping a property the registry claims is a bug.
        (bug,) = corpus.skipped_verdicts([call(result="warned", detail="align")], regs)
        self.assertEqual(bug[1].kind, "bug")

    def test_failure_after_skipping_is_a_gap_not_a_bug(self):
        # Record 0271: a warn-mode add whose data came only from skipped keys
        # fails with `missing_property`; the skipped keys are the gap.
        (row,) = corpus.parse(
            "@ word/x.sh\n3\tadd\tdocx\tchart\tmissing_property\t"
            "data;skipped:style,colour,+2\t-\n")
        self.assertEqual((row.detail, row.after_skip), ("data", "style,colour,+2"))
        self.assertTrue(row.refused and row.failed_after_skipping)
        self.assertEqual(row.skipped(), ["style", "colour", "+2"])
        regs = {"docx": registry([PARAGRAPH, CHART])}
        (verdict,) = corpus.classify([row], regs)
        self.assertEqual((verdict.kind, verdict.rule), ("gap", "F3b"))
        skipped = corpus.skipped_verdicts(
            [call(result="missing_property", detail="data", after_skip="style")], regs)
        self.assertEqual([v.rule for _, v in skipped], ["docx|add|paragraph|skipped|style"])
        # Skipping a property the registry claims is still a bug ...
        (bug,) = corpus.skipped_verdicts(
            [call(result="missing_property", detail="data", after_skip="align")], regs)
        self.assertEqual(bug[1].kind, "bug")
        # ... the same failure without skipped properties stays untriaged ...
        (plain,) = self.classify([call(kind="chart", result="missing_property", detail="data")])
        self.assertEqual(plain.kind, "bug")
        # ... and so does another failure after skipping: only a missing
        # property is explained by the skip.
        (other,) = self.classify([call(result="invalid_value", detail="size", after_skip="x")])
        self.assertEqual(other.kind, "bug")
        # The call's own failure is classified first (here a recorded delta).
        (delta,) = self.classify([call(verb="batch", fmt="pptx", kind="", result="missing_property",
                                       detail="layout", after_skip="transition")])
        self.assertEqual((delta.kind, delta.rule), ("delta", "D4b"))
        # A later path into the failed add is a cascade.
        after = self.classify([row, call(verb="set", result="path_not_found", target="p")])
        self.assertEqual(after[1].kind, "cascade")

    def test_batch_slide_without_layout_is_delta_d4b(self):
        (verdict,) = self.classify([
            call(verb="batch", fmt="pptx", kind="", result="missing_property", detail="layout"),
        ])
        self.assertEqual((verdict.kind, verdict.rule), ("delta", "D4b"))

    def test_fatal_results_are_bugs(self):
        (verdict,) = self.classify([call(result="internal_error")])
        self.assertEqual((verdict.kind, verdict.rule), ("bug", "fatal"))

    def test_unclaimed_pass_is_reported(self):
        regs = {"docx": registry([PARAGRAPH, CHART])}
        found = corpus.unclaimed([call(kind="chart"), call()], regs)
        self.assertEqual(len(found), 1)
        self.assertIn("no add on `chart`", found[0][1])


class GenerateTests(unittest.TestCase):
    def test_committed_table_is_current(self):
        self.assertEqual(corpus.main(["--check"]), 0)

    def test_render_is_deterministic_and_counts(self):
        calls = [call(), call(line=0, verb="view"), call(line=2, result="usage", detail="A,B")]
        verdicts = corpus.classify(calls, {"docx": registry([PARAGRAPH])})
        first = corpus.render(calls, verdicts, "abc")
        self.assertEqual(first, corpus.render(calls, verdicts, "abc"))
        self.assertIn("| word | 2 | 1 | 0 | 0 | 0 | 1 | 0 | 0 | 0 | 1/1 |", first)

    def test_check_detects_a_stale_table(self):
        with tempfile.TemporaryDirectory() as tmp:
            stale = Path(tmp) / "table.md"
            stale.write_text("stale", encoding="utf-8")
            self.assertEqual(corpus.main(["--check", "--output", str(stale)]), 1)


if __name__ == "__main__":
    unittest.main()
