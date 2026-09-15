"""Source-level guardrails for the split benchmark entry points.

These checks intentionally do not invoke Cargo. They keep the normal latency
binary forbid-safe and make it difficult to move the global allocator back into
the shared harness by accident.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PERF_BASELINE = ROOT / "tools" / "perf-baseline"


class PerfBaselineSourcePolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = (PERF_BASELINE / "Cargo.toml").read_text(encoding="utf-8")
        cls.normal = (PERF_BASELINE / "src" / "main.rs").read_text(encoding="utf-8")
        cls.library = (PERF_BASELINE / "src" / "lib.rs").read_text(encoding="utf-8")
        cls.allocator = (
            PERF_BASELINE / "src" / "bin" / "litchi-perf-baseline-alloc.rs"
        ).read_text(encoding="utf-8")
        cls.allocator_support = (
            PERF_BASELINE / "src" / "bin" / "support" / "counting_allocator.rs"
        ).read_text(encoding="utf-8")
        cls.xlsb_crud = (
            PERF_BASELINE / "src" / "bin" / "xlsb_crud.rs"
        ).read_text(encoding="utf-8")
        cls.metrics = (
            PERF_BASELINE / "src" / "allocation_metrics.rs"
        ).read_text(encoding="utf-8")
        cls.filesystem = (PERF_BASELINE / "src" / "filesystem.rs").read_text(
            encoding="utf-8"
        )
        cls.producer_shape = (PERF_BASELINE / "src" / "producer_shape.rs").read_text(
            encoding="utf-8"
        )
        cls.ole2_range_source = (
            PERF_BASELINE / "src" / "ole2_range_source.rs"
        ).read_text(encoding="utf-8")
        cls.workflow = (ROOT / ".github" / "workflows" / "perf-baseline.yml").read_text(
            encoding="utf-8"
        )

    def test_normal_entry_is_unconditionally_forbid_safe(self):
        self.assertIn("#![forbid(unsafe_code)]", self.normal)
        self.assertNotIn("cfg_attr", self.normal)
        self.assertNotRegex(self.normal, r"\bunsafe\b")
        self.assertNotIn("allocator-metrics", self.normal)
        self.assertNotIn("allocation_metrics::enable", self.normal)

    def test_shared_harness_library_contains_no_unsafe_allocator_surface(self):
        self.assertIn("#![forbid(unsafe_code)]", self.library)
        self.assertNotIn("unsafe impl", self.library)
        self.assertNotIn("unsafe fn", self.library)
        self.assertNotIn("#[global_allocator]", self.library)
        self.assertNotIn("GlobalAlloc", self.library)

    def test_allocator_target_owns_global_allocator_unsafe_surface(self):
        self.assertIn('#[path = "support/counting_allocator.rs"]\nmod allocator;', self.allocator)
        self.assertIn("litchi_perf_baseline::allocation_metrics::enable();", self.allocator)
        self.assertNotIn("counting_allocator", self.normal)
        self.assertNotIn("counting_allocator", self.library)
        self.assertIn("use litchi_perf_baseline::allocation_metrics;", self.allocator_support)
        self.assertNotIn("use super::allocation_metrics;", self.allocator_support)
        self.assertIn("unsafe impl GlobalAlloc", self.allocator_support)
        self.assertIn("#[global_allocator]", self.allocator_support)
        self.assertIn("System.alloc", self.allocator_support)
        self.assertNotIn("#![forbid(unsafe_code)]", self.allocator)
        self.assertNotIn("unsafe impl", self.metrics)
        self.assertNotIn("#[global_allocator]", self.metrics)

    def test_manifest_points_feature_target_at_distinct_source(self):
        self.assertIn(
            'path = "src/bin/litchi-perf-baseline-alloc.rs"',
            self.manifest,
        )
        self.assertIn('required-features = ["allocator-metrics"]', self.manifest)
        self.assertNotIn(
            'path = "src/main.rs"\nrequired-features = ["allocator-metrics"]',
            self.manifest,
        )

    def test_xlsb_crud_target_is_opt_in_and_consumes_timed_outcomes(self):
        self.assertIn('xlsb-crud = ["litchi/xlsb"]', self.manifest)
        self.assertIn('path = "src/bin/xlsb_crud.rs"', self.manifest)
        self.assertIn('required-features = ["xlsb-crud"]', self.manifest)
        self.assertIn(
            'litchi-xlsb = { path = "../../crates/litchi-xlsb" }',
            self.manifest,
        )
        self.assertNotRegex(
            self.manifest,
            r'features = \[[^\]]*"xlsb"[^\]]*\]\s*\}',
        )
        self.assertIn("let outcome = std::hint::black_box(outcome);", self.xlsb_crud)
        self.assertIn(
            "binary_identity: litchi_perf_baseline::BinaryIdentity",
            self.xlsb_crud,
        )
        self.assertIn(
            "litchi_perf_baseline::current_executable_identity()?",
            self.xlsb_crud,
        )

    def test_allocator_target_is_built_and_tested_by_ci(self):
        self.assertIn("--features allocator-metrics", self.workflow)
        self.assertIn("--bin litchi-perf-baseline-alloc", self.workflow)
        self.assertRegex(self.workflow, r"cargo check[\s\S]+allocator-metrics")
        self.assertRegex(self.workflow, r"cargo test[\s\S]+allocator-metrics")

    def test_real_producer_security_corpus_is_locked_and_after_normal_harness(self):
        self.assertIn("permissions:\n  contents: read\n  actions: read", self.workflow)
        self.assertIn(
            'cargo test --locked --manifest-path "$PERF_MANIFEST" '
            "--lib security_corpus -- --ignored",
            self.workflow,
        )
        normal = self.workflow.index("- name: Check standalone harness")
        security = self.workflow.index(
            "- name: Run real-producer security correctness corpus"
        )
        allocator = self.workflow.index("- name: Check allocator-only benchmark target")
        self.assertLess(normal, security)
        self.assertLess(security, allocator)

    def test_allocator_target_is_executed_and_compared_by_ci(self):
        self.assertRegex(
            self.workflow,
            r"cargo run --release[\s\S]+--features allocator-metrics[\s\S]+"
            r"--bin litchi-perf-baseline-alloc",
        )
        for argument in (
            "--warmup 3",
            "--samples 15",
            "--case opc_file_eager_open",
            "--filesystem-cache warm,cold-requested",
            '--json "$LITCHI_SMOKE_DIR/current.json"',
        ):
            self.assertIn(argument, self.workflow)
        self.assertIn("LITCHI_SMOKE_DIR: target/perf/allocator-smoke", self.workflow)
        self.assertIn("tools/perf_compare.py", self.workflow)
        self.assertIn(
            "LITCHI_ALLOCATOR_POLICY: "
            "docs/performance/perf-regression-policy-allocator-v1.json",
            self.workflow,
        )
        # Change 0626 moved the comparison's expected shape out of an inline
        # workflow heredoc and into the checked smoke policy, where
        # tools/perf_smoke_baseline.py reads it and a unit suite covers it.
        import json

        smoke_policy = json.loads(
            (ROOT / "docs/performance/perf-smoke-baseline-policy-v1.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertIn(
            "LITCHI_SMOKE_POLICY: "
            "docs/performance/perf-smoke-baseline-policy-v1.json",
            self.workflow,
        )
        self.assertEqual(
            smoke_policy["comparator_policy"],
            "docs/performance/perf-regression-policy-allocator-v1.json",
        )
        expectations = smoke_policy["self_comparison_expectations"]
        self.assertEqual(expectations["latency_claims"], "withheld_instrumentation")
        self.assertEqual(expectations["compared_metrics"], 20)
        self.assertEqual(expectations["matched_results"], 2)
        self.assertEqual(expectations["regressions"], 0)

    def test_allocator_manifest_selects_filesystem_case_and_pinned_corpus(self):
        import json

        policy = json.loads(
            (ROOT / "docs/performance/perf-regression-policy-allocator-v1.json").read_text(
                encoding="utf-8"
            )
        )
        manifest = json.loads(
            (
                ROOT
                / "docs/performance/results/perf-regression-allocator-manifest-v1.json"
            ).read_text(encoding="utf-8")
        )
        self.assertEqual(policy["required_cases"], ["opc_file_eager_open"])
        self.assertEqual(policy["required_cases"], manifest["required_cases"])
        self.assertEqual(policy["expected_result_count"], manifest["result_count"])
        self.assertEqual(
            policy["expected_result_keys_sha256"], manifest["result_keys_sha256"]
        )
        self.assertEqual(policy["result_key_fields"], manifest["result_key_fields"])
        self.assertEqual(policy["metric_classes"][0]["presence"], "required")
        self.assertEqual(len(policy["metric_classes"]), 1)
        for field in (
            "allocation_calls",
            "deallocation_calls",
            "reallocation_calls",
            "failed_allocation_calls",
            "allocated_bytes",
            "deallocated_bytes",
            "live_bytes_before",
            "live_bytes_after",
            "peak_live_bytes_before",
            "peak_live_bytes_after",
        ):
            self.assertIn(
                f"operation_metrics/allocation/{field}/values",
                policy["metric_classes"][0]["path_globs"],
            )
        self.assertNotIn("rss", json.dumps(policy["metric_classes"]))
        self.assertNotIn("work", json.dumps(policy["metric_classes"]))
        for unrelated in ("copied_bytes", "decompressed_bytes", "recompressed_bytes"):
            self.assertNotIn(unrelated, json.dumps(policy["metric_classes"]))

    def test_producer_shape_module_owns_no_unsafe_or_ambient_surface(self):
        self.assertIn("mod producer_shape;", self.library)
        self.assertNotIn("unsafe", self.producer_shape)
        self.assertNotIn("#[global_allocator]", self.producer_shape)
        self.assertNotIn("std::env", self.producer_shape)
        self.assertNotIn("Command", self.producer_shape)

    def test_real_file_selectors_are_opt_in_bounded_and_self_identifying(self):
        # `--real-file` is the only input in this harness whose bytes come from
        # outside the process. Keep it bounded, keep its identity in the
        # corpus, and keep it out of the default matrix.
        self.assertIn(
            "const MAX_REAL_FILE_BYTES: u64 = 32 * 1024 * 1024;", self.producer_shape
        )
        self.assertIn(
            "fn build_xlsx_real_file_corpus(path: &Path)", self.producer_shape
        )
        self.assertIn("struct RealFileProvenance", self.producer_shape)
        for field in ("path: String", "bytes: u64", "sha256: String"):
            self.assertIn(field, self.producer_shape)
        # Exactly one whole-file read, and it is the bounded one.
        self.assertEqual(self.producer_shape.count("fs::read("), 1)
        self.assertIn('"--real-file" => {', self.library)
        self.assertIn("--real-file PATH", self.library)

    def test_producer_shape_selectors_are_absent_from_the_default_matrix(self):
        start = self.library.index("const DEFAULT: ")
        end = self.library.index("];", start)
        default_matrix = self.library[start:end]
        for case in (
            "XlsxProducerMediumSourceOpen",
            "XlsxProducerMediumSourceSelectedCell",
            "XlsxProducerMediumSourcePlanning",
            "XlsxProducerMediumSourceOneEditSave",
            "XlsxProducerDenseSourceOpen",
            "XlsxProducerDenseSourceSelectedCell",
            "XlsxProducerDenseSourcePlanning",
            "XlsxProducerDenseSourceOneEditSave",
            "XlsxProducerMediumControlSelectedCell",
            "XlsxProducerMediumControlPlanning",
            "XlsxProducerDenseControlSelectedCell",
            "XlsxProducerDenseControlPlanning",
            "XlsxRealFileSourceOpen",
            "XlsxRealFileSourceSelectedCell",
            "DocxProducerSourceSelectedParagraph",
            "PptxProducerSourceSelectedSlide",
        ):
            self.assertIn(f"Self::{case}", self.library)
            self.assertNotIn(case, default_matrix)

    def test_ole2_range_source_module_owns_no_unsafe_or_ambient_surface(self):
        self.assertIn("mod ole2_range_source;", self.library)
        self.assertNotIn("unsafe", self.ole2_range_source)
        self.assertNotIn("#[global_allocator]", self.ole2_range_source)
        self.assertNotIn("std::env", self.ole2_range_source)
        self.assertNotIn("Command", self.ole2_range_source)

    def test_ole2_file_selectors_are_opt_in_bounded_and_self_identifying(self):
        # `--ole2-file` is the second input whose bytes come from outside the
        # process. Keep it bounded, keep its identity in the corpus, and keep
        # it out of the default matrix.
        self.assertIn(
            "const MAX_OLE2_FILE_BYTES: u64 = 32 * 1024 * 1024;",
            self.ole2_range_source,
        )
        self.assertIn("fn build_xls_corpus(path: &Path)", self.ole2_range_source)
        self.assertIn("fn build_ppt_corpus(path: &Path)", self.ole2_range_source)
        # Exactly one whole-file read, and it is the bounded one.
        self.assertEqual(self.ole2_range_source.count("fs::read("), 1)
        self.assertIn('"--ole2-file" => {', self.library)
        self.assertIn("--ole2-file PATH", self.library)

    def test_ole2_range_source_selectors_are_absent_from_the_default_matrix(self):
        start = self.library.index("const DEFAULT: ")
        end = self.library.index("];", start)
        default_matrix = self.library[start:end]
        for case in (
            "XlsRangeSourceOpen",
            "XlsRangeSourceOpenListWorksheets",
            "XlsRangeSourceOpenOneCell",
            "XlsRangeSourceOpenAllCells",
            "XlsRangeSourceOpenFullText",
            "XlsOwnedSourceControlOpen",
            "XlsOwnedSourceControlOpenListWorksheets",
            "XlsOwnedSourceControlOpenOneCell",
            "XlsOwnedSourceControlOpenAllCells",
            "XlsOwnedSourceControlOpenFullText",
            "PptRangeSourceOpen",
            "PptRangeSourceOpenOneShapeText",
            "PptOwnedSourceControlOpen",
            "PptOwnedSourceControlOpenOneShapeText",
        ):
            self.assertIn(f"Self::{case}", self.library)
            self.assertNotIn(case, default_matrix)

    def test_region_scope_is_static_and_heap_free(self):
        self.assertIn("pub scope: Scope", self.metrics)
        self.assertIn("const SCOPE: Scope", self.metrics)
        self.assertNotIn("String", self.metrics)
        self.assertNotIn("to_owned", self.metrics)
        self.assertRegex(
            self.metrics,
            re.compile(r"RegionState::Unavailable\s*=>\s*Some\(Sample::unavailable\(\)\)"),
        )

    def test_child_error_path_finishes_region_before_propagating_operation_error(self):
        finish = self.filesystem.index("let allocation_metrics = allocation_region.finish();")
        propagate = self.filesystem.index("let counter = counter_result?;")
        self.assertLess(finish, propagate)
        self.assertIn("allocation_metrics,", self.filesystem[propagate:])


if __name__ == "__main__":
    unittest.main()
