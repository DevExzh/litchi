from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools import officecli_capability_delta as capability_delta
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
            call(verb="watch", result="usage", detail="watch"),
            call(result="usage", detail="A,Product"),
        ])
        self.assertEqual([v.rule for v in verdicts], ["G1", "D7"])

    def test_warned_calls_and_their_skipped_properties(self):
        calls = [call(result="warned", detail="style,colour,font.{bold,size}")]
        regs = {"docx": registry([PARAGRAPH, CHART])}
        (verdict,) = corpus.classify(calls, regs)
        self.assertEqual((verdict.kind, verdict.rule), ("warned", "F3"))
        # A warned call is not a refusal, so it starts no cascade.
        self.assertFalse(calls[0].refused)
        skipped = corpus.skipped_verdicts(calls, regs)
        self.assertEqual([v.rule for _, v in skipped], [
            "docx|add|paragraph|warned|style",
            "docx|add|paragraph|warned|colour",
            "docx|add|paragraph|warned|font.bold",
            "docx|add|paragraph|warned|font.size",
        ])
        # Skipping a property the registry claims is a bug.
        (bug,) = corpus.skipped_verdicts([call(result="warned", detail="align")], regs)
        self.assertEqual(bug[1].kind, "bug")

    def test_skipped_lists_are_complete(self):
        # Record 0339: the outcome lists every skipped name; a `head.{...}`
        # group stands for consecutive names sharing a first segment.
        self.assertEqual(corpus.expand_skipped("title.{font,size},fill,a.{b.c,d}"),
                         ["title.font", "title.size", "fill", "a.b.c", "a.d"])
        self.assertEqual(corpus.expand_skipped(""), [])
        # A list truncated by the old 48-character bound is refused.
        with self.assertRaises(SystemExit):
            corpus.expand_skipped("style,+2", "word/x.sh:3")

    def test_failure_after_skipping_is_a_gap_not_a_bug(self):
        # Record 0271: a warn-mode add whose data came only from skipped keys
        # fails with `missing_property`; the skipped keys are the gap.
        (row,) = corpus.parse(
            "@ word/x.sh\n3\tadd\tdocx\tchart\tmissing_property\t"
            "data;skipped:style,series1.{name,values}\t-\n")
        self.assertEqual((row.detail, row.after_skip), ("data", "style,series1.{name,values}"))
        self.assertTrue(row.refused and row.failed_after_skipping)
        self.assertEqual(row.skipped(), ["style", "series1.name", "series1.values"])
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


def classes(entries: list[dict], scripts: list[dict] | None = None, deltas: list | None = None,
            tiers: list[dict] | None = None) -> corpus.CorpusDispositions:
    loaded = capability_delta.Dispositions({
        "tiers": tiers if tiers is not None else [
            {"tier": "P0", "keys": ["docx:paragraph", "docx:chart"], "basis": "b"},
            {"tier": "P2", "keys": ["corpus:G4"], "basis": "b"},
            {"tier": "P1", "keys": ["corpus:G1|*"], "basis": "b"},
        ],
        "dispositions": entries,
        "corpusScripts": scripts if scripts is not None else [
            {"keys": ["word/*"], "tier": "P0", "basis": "b"},
            {"keys": ["word/later.sh"], "tier": "P2", "basis": "b"},
        ],
    }, "test")
    oc = capability_delta.OcElement("paragraph", "paragraph", ["p"], [], {},
                                    {"align": {"type": "enum", "aliases": ["alignment"], "set": True}})
    regs = {"docx": corpus.Registry("docx", {"elements": [PARAGRAPH, CHART], "gaps": []}, [oc])}
    return corpus.CorpusDispositions(loaded, deltas or [], regs)


def rule(id: str, disposition: str, **extra) -> dict:
    return {"id": id.lower(), "keys": [f"corpus:{id}"], "disposition": disposition, "reason": "r", **extra}


class DispositionTests(unittest.TestCase):
    def test_new_verbs_are_not_missing_verbs(self):
        verdicts = corpus.classify([call(verb="import", result="usage", detail="A,B"),
                                    call(verb="watch", result="usage", detail="watch")],
                                   {"docx": registry([PARAGRAPH])})
        self.assertEqual([v.rule for v in verdicts], ["D7", "G1"])
        self.assertNotIn("`import`, `merge`", corpus.RULES_BY_ID["G1"].officecli)

    def test_cascades_record_their_root(self):
        calls = [call(), call(kind="chart", result="unsupported_element"), call(verb="get", result="ok"),
                 call(verb="set", result="path_not_found", target="p")]
        verdicts = corpus.classify(calls, {"docx": registry([PARAGRAPH, CHART])})
        self.assertEqual((verdicts[3].kind, verdicts[3].root), ("cascade", 1))

    def test_groups_inherit_the_delta_and_named_classes_resolve(self):
        prop = ("paragraph", "align", "enum", "-s-", "missing", "", "", "gap")
        delta = capability_delta.FormatDelta("docx", properties=[prop])
        delta.element_dispositions = []
        delta.property_dispositions = [capability_delta.Resolution("deferred-P1", "P1", "0332", "breadth")]
        found = classes([rule("G1|validate", "planned", ref="WP5"),
                         {"id": "own", "keys": ["corpus:docx|*|paragraph|*|bold"], "disposition": "planned",
                          "ref": "WP4"},
                         {"id": "fallback", "keys": ["corpus:docx|*|*|*|*"], "disposition": "deferred",
                          "ref": "0332"}],
                        deltas=[delta])
        inherited = found.group("docx|set|paragraph|warned|alignment")
        self.assertEqual((inherited.entry, inherited.via), ("breadth", "delta `paragraph.align`"))
        self.assertEqual(found.group("docx|set|paragraph|warned|bold").entry, "own")
        self.assertEqual(found.group("docx|set|paragraph|warned|shade").disposition, "deferred-P1")
        g1 = corpus.RULES_BY_ID["G1"]
        self.assertEqual(found.rule(g1, call(verb="validate")).ref, "WP5")
        self.assertIsNone(found.rule(g1, call(verb="raw-set")).disposition)

    def test_problems(self):
        calls = [
            call(verb="watch", result="usage", detail="watch"),
            call(verb="open", result="skip"),
            call(result="warned", detail="shade,(x)"),
            call(kind="chart", result="unsupported_element"),
            call(verb="remove", result="invalid_argument", script="word/later.sh"),
            call(script="excel/x.sh"),
        ]
        found = classes([rule("G1|validate", "planned", ref="WP5"), rule("D5", "planned", ref="WP5"),
                         rule("G4", "deferred-P1"), rule("Z9", "wontfix"),
                         {"id": "p0", "keys": ["corpus:docx|add|paragraph|warned|*"], "disposition": "deferred",
                          "ref": "0332"}],
                        tiers=[{"tier": "P0", "keys": ["docx:paragraph", "docx:chart", "docx:paragraph.shade"],
                                "basis": "b"},
                               {"tier": "P2", "keys": ["corpus:G4"], "basis": "b"},
                               {"tier": "P1", "keys": ["corpus:G1|*"], "basis": "b"}])
        verdicts = corpus.classify(calls, found.registries)
        problems = corpus.corpus_problems(calls, verdicts, found, found.registries)
        self.assertIn("corpus D5: a recorded delta is wontfix or out-of-scope, not `planned`", problems)
        self.assertIn("corpus G4: `deferred-P1` on a P2 row (g4)", problems)
        self.assertIn("corpus docx|add|paragraph|warned|shade: a P0 row is `deferred` (p0); P0 is supported or planned",
                      problems)
        self.assertIn("corpus docx|add|chart|unsupported_element|: no disposition", problems)
        self.assertIn("corpus script excel/x.sh: no tier", problems)
        self.assertIn("stale disposition key `corpus:Z9`: it classifies no corpus class", problems)

    def test_the_v1_subset(self):
        calls = [
            call(),                                                  # passed
            call(verb="watch", result="usage", detail="watch"),      # G1: planned, open
            call(verb="remove", result="invalid_argument"),          # G4: deferred, accounted
            call(result="warned", detail="style,shade"),             # shade: no disposition
            call(verb="mark", result="usage", script="word/later.sh"),  # outside the subset
            call(kind="chart", result="unsupported_element"),        # no disposition: open
            call(verb="set", result="path_not_found", target="p"),   # cascade of the chart add
        ]
        found = classes([rule("G1|watch", "planned", ref="WP5"), rule("G1|mark", "planned", ref="WP5"),
                         rule("G4", "deferred"),
                         {"id": "style", "keys": ["corpus:docx|add|paragraph|warned|style"],
                          "disposition": "wontfix", "reason": "r"}])
        verdicts = corpus.classify(calls, found.registries)
        subset = corpus.v1_subset(calls, verdicts, found, found.registries)
        self.assertFalse(subset.green)
        states = [(calls[index].verb, state) for index, state, _ in subset.lines]
        self.assertEqual(states, [("add", "passed"), ("watch", "open"), ("remove", "accounted"),
                                  ("add", "open"), ("add", "open"), ("set", "open")])
        blockers = [corpus.blocker(found_classes) for _, state, found_classes in subset.lines if state == "open"]
        self.assertEqual(blockers[0], "planned WP5")
        self.assertEqual(blockers[1], "no disposition (`docx|add|paragraph|warned|shade`)")
        self.assertEqual(blockers[2], "no disposition (`docx|add|chart|unsupported_element|`)")
        self.assertEqual(blockers[3], blockers[2])  # the cascade carries its root's class
        rendered = corpus.render(calls, verdicts, "abc", classes=found, subset=subset)
        self.assertIn("**Status: red.** 6 v1 lines: 1 passed, 1 accounted, 4 open.", rendered)
        self.assertIn("| planned WP5 | 1 |", rendered)
        green = corpus.v1_subset(calls[:3:2], corpus.classify(calls[:3:2], found.registries), found,
                                 found.registries)
        self.assertTrue(green.green)

    def test_check_fails_on_disposition_problems(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dispositions.json"
            path.write_text('{"tiers": [], "dispositions": []}', encoding="utf-8")
            self.assertEqual(corpus.main(["--check", "--dispositions", str(path)]), 1)


if __name__ == "__main__":
    unittest.main()
