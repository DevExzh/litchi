import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from tools import perf_decision


REPO_ROOT = Path(__file__).resolve().parents[1]
AUDIT_SHA256 = "2712e39c456d7528c4c9347d6f085dd33253bf43d389d265b979ffee606b29ea"
AUDIT_PATH = (
    "/home/zhuhe/litchi-performance-0967/review/"
    "independent-raw-audit-0967-v2.json"
)


def _facet(status, *reason_codes):
    return perf_decision.Facet(status, tuple(reason_codes))


def _facts(
    *,
    completeness=perf_decision.CompletenessStatus.COMPLETE,
    custody=perf_decision.FacetStatus.PASS,
    correctness=perf_decision.FacetStatus.PASS,
    readiness=perf_decision.FacetStatus.PASS,
    performance=perf_decision.FacetStatus.PASS,
    benefit=perf_decision.FacetStatus.PASS,
    performance_reasons=None,
):
    completeness_reason = (
        "terminal_packets_complete"
        if completeness is perf_decision.CompletenessStatus.COMPLETE
        else "terminal_packets_missing"
    )
    benefit_reason = (
        "predeclared_benefit_pass"
        if benefit is perf_decision.FacetStatus.PASS
        else "predeclared_benefit_missing"
        if benefit is perf_decision.FacetStatus.MISSING
        else "predeclared_benefit_failed"
    )
    return perf_decision.QualificationFacts(
        candidate="0967-doc-sprm-vector-free",
        source_revision="893072657ad8b1d726bfce351b5ec2e98f084e94",
        completeness=perf_decision.Completeness(
            completeness,
            (completeness_reason,),
        ),
        custody=_facet(custody, "custody_and_schema_complete"),
        correctness=_facet(correctness, "semantic_controls_pass"),
        readiness=_facet(readiness, "aa_readiness"),
        performance=_facet(
            performance,
            *(performance_reasons or ("performance_gates_pass",)),
        ),
        benefit=_facet(benefit, benefit_reason),
        provenance=(
            perf_decision.ProvenanceRef(
                kind="independent_raw_audit",
                path=AUDIT_PATH,
                sha256=AUDIT_SHA256,
            ),
        ),
    )


def _input_mapping(facts):
    return {
        "schema": perf_decision.INPUT_SCHEMA,
        "candidate": facts.candidate,
        "source_revision": facts.source_revision,
        "completeness": facts.completeness.as_dict(),
        "facets": {
            name: facts.facets()[name].as_dict()
            for name in ("custody", "correctness", "readiness", "performance", "benefit")
        },
        "provenance": [item.as_dict() for item in facts.provenance],
    }


class DecisionTests(unittest.TestCase):
    def test_normalized_0967_facts_reject_complete_failed_performance_gate(self):
        facts = _facts(
            performance=perf_decision.FacetStatus.FAIL,
            performance_reasons=(
                "native_noninferiority_failed",
                "allocator_memory_rss_guard_failed",
                "timed_allocator_guard_failed",
            ),
        )
        # The normalized fixture mirrors v2: complete custody/semantic/AA
        # evidence and an interpretable performance rejection.
        result = perf_decision.decide(facts)

        self.assertEqual(result.decision, perf_decision.DecisionCode.REJECT_CANDIDATE)
        self.assertIs(result.candidate_adoption, False)
        self.assertIn("performance:native_noninferiority_failed", result.reasons)
        self.assertIn("performance:allocator_memory_rss_guard_failed", result.reasons)
        self.assertNotIn("incomplete", result.reasons)
        self.assertEqual(result.as_dict()["provenance"][0]["sha256"], AUDIT_SHA256)

    def test_missing_benefit_is_incomplete_without_candidate_verdict(self):
        facts = _facts(benefit=perf_decision.FacetStatus.MISSING)
        result = perf_decision.decide(facts)

        self.assertEqual(result.decision, perf_decision.DecisionCode.INCOMPLETE)
        self.assertIsNone(result.candidate_adoption)
        self.assertIn("benefit:predeclared_benefit_missing", result.reasons)

    def test_readiness_failure_is_uninterpretable_even_with_performance_failure(self):
        facts = _facts(
            readiness=perf_decision.FacetStatus.FAIL,
            performance=perf_decision.FacetStatus.FAIL,
        )
        result = perf_decision.decide(facts)

        self.assertEqual(result.decision, perf_decision.DecisionCode.UNINTERPRETABLE)
        self.assertIsNone(result.candidate_adoption)
        self.assertIn("readiness:aa_readiness", result.reasons)
        self.assertNotIn("performance:performance_gates_pass", result.reasons)

    def test_custody_failure_is_uninterpretable(self):
        facts = _facts(custody=perf_decision.FacetStatus.FAIL)
        result = perf_decision.decide(facts)

        self.assertEqual(result.decision, perf_decision.DecisionCode.UNINTERPRETABLE)
        self.assertIsNone(result.candidate_adoption)
        self.assertIn("custody:custody_and_schema_complete", result.reasons)

    def test_all_pass_with_benefit_adopts(self):
        result = perf_decision.decide(_facts())

        self.assertEqual(result.decision, perf_decision.DecisionCode.ADOPT)
        self.assertIs(result.candidate_adoption, True)
        self.assertEqual(result.reasons, ("all_required_gates_passed",))

    def test_no_benefit_rejects_complete_interpretable_evidence(self):
        facts = _facts(benefit=perf_decision.FacetStatus.FAIL)
        result = perf_decision.decide(facts)

        self.assertEqual(result.decision, perf_decision.DecisionCode.REJECT_CANDIDATE)
        self.assertIs(result.candidate_adoption, False)
        self.assertIn("benefit:predeclared_benefit_failed", result.reasons)

    def test_strict_input_rejects_unknown_keys_and_empty_reasons(self):
        mapping = _input_mapping(_facts())
        mapping["unexpected"] = True
        with self.assertRaises(perf_decision.DecisionInputError):
            perf_decision.facts_from_mapping(mapping)

        mapping = _input_mapping(_facts())
        mapping["facets"]["performance"]["reason_codes"] = []
        with self.assertRaises(perf_decision.DecisionInputError):
            perf_decision.facts_from_mapping(mapping)

    def test_strict_json_rejects_duplicate_keys_and_nonfinite_numbers(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.json"
            path.write_text(
                '{"schema":"x","schema":"y"}', encoding="utf-8"
            )
            with self.assertRaises(perf_decision.DecisionInputError):
                perf_decision.load_facts(path)
            path.write_text('{"value":NaN}', encoding="utf-8")
            with self.assertRaises(perf_decision.DecisionInputError):
                perf_decision.load_facts(path)


class DecisionCliTests(unittest.TestCase):
    def test_cli_writes_explicit_fresh_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "input.json"
            output_path = root / "output.json"
            input_path.write_text(
                json.dumps(_input_mapping(_facts()), sort_keys=True), encoding="utf-8"
            )

            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                status = perf_decision.main(
                    ["--input", str(input_path), "--json-out", str(output_path)]
                )

            self.assertEqual(status, 0)
            self.assertEqual(
                json.loads(stdout.getvalue())["decision"], "adopt"
            )
            output = json.loads(output_path.read_text(encoding="utf-8"))
            self.assertEqual(output["schema"], perf_decision.OUTPUT_SCHEMA)
            self.assertEqual(output["candidate_adoption"], True)

    def test_cli_refuses_existing_output_and_preserves_it(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "input.json"
            output_path = root / "output.json"
            input_path.write_text(
                json.dumps(_input_mapping(_facts()), sort_keys=True), encoding="utf-8"
            )
            output_path.write_text("sentinel\n", encoding="utf-8")

            with contextlib.redirect_stderr(io.StringIO()):
                status = perf_decision.main(
                    ["--input", str(input_path), "--json-out", str(output_path)]
                )

            self.assertEqual(status, 2)
            self.assertEqual(output_path.read_text(encoding="utf-8"), "sentinel\n")

    def test_cli_refuses_output_symlink(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "input.json"
            target_path = root / "target.json"
            output_path = root / "output.json"
            input_path.write_text(
                json.dumps(_input_mapping(_facts()), sort_keys=True), encoding="utf-8"
            )
            output_path.symlink_to(target_path)

            with contextlib.redirect_stderr(io.StringIO()):
                status = perf_decision.main(
                    ["--input", str(input_path), "--json-out", str(output_path)]
                )

            self.assertEqual(status, 2)
            self.assertFalse(target_path.exists())
            self.assertTrue(output_path.is_symlink())

    def test_cli_input_output_collision_is_rejected_without_rewrite(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.json"
            original = json.dumps(_input_mapping(_facts()), sort_keys=True)
            path.write_text(original, encoding="utf-8")

            with contextlib.redirect_stderr(io.StringIO()):
                status = perf_decision.main(
                    ["--input", str(path), "--json-out", str(path)]
                )

            self.assertEqual(status, 2)
            self.assertEqual(path.read_text(encoding="utf-8"), original)

    def test_help_and_invalid_arguments_do_not_create_output(self):
        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "output.json"
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(
                io.StringIO()
            ), self.assertRaises(SystemExit) as help_exit:
                perf_decision.main(["--help"])
            self.assertEqual(help_exit.exception.code, 0)
            self.assertFalse(output_path.exists())

            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(
                io.StringIO()
            ), self.assertRaises(SystemExit) as argument_exit:
                perf_decision.main(["--unknown", "value"])
            self.assertEqual(argument_exit.exception.code, 2)
            self.assertFalse(output_path.exists())

    def test_cli_malformed_or_wrong_schema_input_has_no_output_side_effect(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "input.json"
            output_path = root / "output.json"
            for payload in ("{not-json\n", json.dumps({"schema": "wrong"})):
                input_path.write_text(payload, encoding="utf-8")
                source_before = input_path.read_bytes()
                completed = subprocess.run(
                    [
                        sys.executable,
                        str(Path(perf_decision.__file__)),
                        "--input",
                        str(input_path),
                        "--json-out",
                        str(output_path),
                    ],
                    cwd=REPO_ROOT,
                    capture_output=True,
                    text=True,
                    check=False,
                )

                self.assertEqual(completed.returncode, 2)
                self.assertFalse(output_path.exists())
                self.assertEqual(input_path.read_bytes(), source_before)


if __name__ == "__main__":
    unittest.main()
