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
        cls.facade_ole2 = (PERF_BASELINE / "src" / "facade_ole2.rs").read_text(
            encoding="utf-8"
        )
        cls.ordinary_save = (PERF_BASELINE / "src" / "ordinary_save.rs").read_text(
            encoding="utf-8"
        )
        cls.marker_shape = (PERF_BASELINE / "src" / "marker_shape.rs").read_text(
            encoding="utf-8"
        )
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


    def test_facade_and_ordinary_save_modules_own_no_unsafe_or_ambient_surface(self):
        # Change 0638's two modules. The facade module reads a caller-named
        # path; the ordinary-save module writes into a caller-named directory
        # through `filesystem::scratch_root`, which is where `std::env` lives
        # in this harness. Neither may reach for an ambient one itself.
        self.assertIn("mod facade_ole2;", self.library)
        self.assertIn("mod ordinary_save;", self.library)
        for module in (self.facade_ole2, self.ordinary_save):
            self.assertNotIn("unsafe", module)
            self.assertNotIn("#[global_allocator]", module)
            self.assertNotIn("std::env", module)
            self.assertNotIn("Command", module)
        self.assertIn(
            "crate::filesystem::scratch_root(requested_root, \"ordinary-save\")",
            self.ordinary_save,
        )

    def test_facade_selectors_reuse_the_bounded_ole2_file_input(self):
        # The facade family names no new caller-supplied input: it reads the
        # same `--ole2-file` fixture change 0627 bounded, through the same
        # bounded reader, and the module performs no whole-file read of its
        # own.
        self.assertEqual(self.facade_ole2.count("fs::read("), 0)
        self.assertIn("ole2_range_source::{cfb_inventory, provenance_of, read_bounded}", self.facade_ole2)
        self.assertIn("fn build_doc_corpus(path: &Path)", self.facade_ole2)
        self.assertIn("fn build_ppt_corpus(path: &Path)", self.facade_ole2)
        self.assertIn("struct FacadeEvidence", self.facade_ole2)
        for field in ("real_file: RealFileProvenance", "cfb_stream_count: usize"):
            self.assertIn(field, self.facade_ole2)
        # `--ole2-file` now classifies DOC as well, and stays the single
        # authority for a caller-named OLE2 fixture.
        self.assertIn("Format::Doc => &mut inputs.doc,", self.ole2_range_source)
        self.assertIn("fn classify_inputs(paths: &[PathBuf])", self.ole2_range_source)

    def test_ooxml_file_selectors_are_opt_in_bounded_and_self_identifying(self):
        # `--ooxml-file` is the third input whose bytes come from outside the
        # process. Keep it bounded, keep its identity in the corpus, and keep
        # it out of the default matrix.
        self.assertIn(
            "const MAX_OOXML_FILE_BYTES: u64 = 32 * 1024 * 1024;",
            self.ordinary_save,
        )
        # Exactly one read of a caller-named path, and it is the bounded one.
        # Every other `fs::read` in the module reads back an artifact this
        # harness itself just published into its own private workspace.
        self.assertIn("fn read_bounded(path: &Path)", self.ordinary_save)
        self.assertEqual(self.ordinary_save.count("fs::read(path)"), 1)
        for readback in (
            "fs::read(&corpus.workspace.destination)",
            "fs::read(&corpus.workspace.alternate)",
        ):
            self.assertIn(readback, self.ordinary_save)
        self.assertIn('"--ooxml-file" => {', self.library)
        self.assertIn("--ooxml-file PATH", self.library)
        self.assertIn("fn classify(path: &Path) -> Result<Format,", self.ordinary_save)

    def test_ordinary_save_reports_its_publication_and_determinism_evidence(self):
        # The record's three load-bearing claims must be structural, not
        # narrative: the atomic interval names its steps, the byte split names
        # its derivation, and the 0625/0631 invariant is proved per corpus.
        self.assertIn("atomic_publication_steps", self.ordinary_save)
        self.assertIn("repeated_cycles_identical", self.ordinary_save)
        self.assertIn("repeated_saves_identical", self.ordinary_save)
        self.assertIn("payload_bytes_identical_to_source", self.ordinary_save)
        self.assertIn("uncompressed_payload_bytes_regenerated", self.ordinary_save)
        self.assertIn("edit_outcomes_identical", self.ordinary_save)

    def test_change_0638_selectors_are_absent_from_the_default_matrix(self):
        start = self.library.index("const DEFAULT: ")
        end = self.library.index("];", start)
        default_matrix = self.library[start:end]
        for case in (
            "DocFacadeFileOpen",
            "DocFacadeFileFullText",
            "DocFacadeFileOneParagraph",
            "PptFacadeFileOpen",
            "PptFacadeFileFullText",
            "PptFacadeFileOneSlideText",
            "DocxOrdinarySaveLifecycle",
            "DocxOrdinarySaveEdit",
            "DocxOrdinarySaveAtomicPublish",
            "DocxOrdinarySaveCountingPublish",
            "DocxRealFileOrdinarySaveLifecycle",
            "DocxRealFileOrdinarySaveEdit",
            "DocxRealFileOrdinarySaveAtomicPublish",
            "DocxRealFileOrdinarySaveCountingPublish",
            "XlsxOrdinarySaveLifecycle",
            "XlsxOrdinarySaveEdit",
            "XlsxOrdinarySaveAtomicPublish",
            "XlsxOrdinarySaveCountingPublish",
            "XlsxRealFileOrdinarySaveLifecycle",
            "XlsxRealFileOrdinarySaveEdit",
            "XlsxRealFileOrdinarySaveAtomicPublish",
            "XlsxRealFileOrdinarySaveCountingPublish",
            "PptxOrdinarySaveLifecycle",
            "PptxOrdinarySaveEdit",
            "PptxOrdinarySaveAtomicPublish",
            "PptxOrdinarySaveCountingPublish",
            "PptxRealFileOrdinarySaveLifecycle",
            "PptxRealFileOrdinarySaveEdit",
            "PptxRealFileOrdinarySaveAtomicPublish",
            "PptxRealFileOrdinarySaveCountingPublish",
        ):
            self.assertIn(f"Self::{case}", self.library)
            self.assertNotIn(case, default_matrix)

    def test_marker_shape_module_owns_no_unsafe_or_ambient_surface(self):
        # Change 0664's module generates its corpora in memory from packages
        # the production writers emit. It reads no file, reaches for no
        # ambient state and owns no allocator.
        self.assertIn("mod marker_shape;", self.library)
        self.assertNotIn("unsafe", self.marker_shape)
        self.assertNotIn("#[global_allocator]", self.marker_shape)
        self.assertNotIn("std::env", self.marker_shape)
        self.assertNotIn("Command", self.marker_shape)
        self.assertNotIn("fs::read(", self.marker_shape)

    def test_marker_shape_states_its_derivation_and_its_control(self):
        # The shape is derived from real tracked fixtures by a retained
        # script, and the control is a same-length namespace substitution so
        # the pair differs only in the codec branch. Both facts are
        # structural, not narrative.
        for fixture in (
            "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx",
            "test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx",
            "test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx",
        ):
            self.assertIn(fixture, self.marker_shape)
        self.assertIn(
            "docs/performance/results/change-0664/scripts/derive_marker_shape.py",
            self.marker_shape,
        )
        self.assertIn("fn prove_control_is_byte_comparable(", self.marker_shape)
        self.assertIn("fn strip_markers(", self.marker_shape)
        self.assertIn("marked_byte_share_basis_points", self.marker_shape)
        self.assertIn("skeleton_marked_member_count", self.marker_shape)
        self.assertIn("sink_refusal", self.marker_shape)

    def test_ordinary_save_reports_allocation_regions(self):
        # Change 0649's harness gap (b): the ordinary-save family emitted no
        # allocation metrics because it never opened a region. It does now,
        # and the region is exactly the timed interval.
        self.assertIn("allocation_metrics::begin()", self.ordinary_save)
        self.assertIn("operation_metrics::InProcessObservation", self.ordinary_save)
        self.assertIn(
            "from_in_process_observations_without_sink", self.ordinary_save
        )
        self.assertEqual(
            self.ordinary_save.count("allocation_metrics::begin()"), 4
        )

    def test_change_0664_selectors_are_absent_from_the_default_matrix(self):
        start = self.library.index("const DEFAULT: ")
        end = self.library.index("];", start)
        default_matrix = self.library[start:end]
        for case in (
            "PptxMarkerEagerFullText",
            "PptxMarkerSourceFullText",
            "PptxMarkerControlEagerFullText",
            "PptxMarkerControlSourceFullText",
            "DocxMarkerEagerFullText",
            "DocxMarkerSourceFullText",
            "DocxMarkerControlEagerFullText",
            "DocxMarkerControlSourceFullText",
            "DocxMarkerOrdinarySaveLifecycle",
            "DocxMarkerOrdinarySaveEdit",
            "DocxMarkerOrdinarySaveAtomicPublish",
            "DocxMarkerOrdinarySaveCountingPublish",
            "DocxMarkerControlOrdinarySaveLifecycle",
            "DocxMarkerControlOrdinarySaveEdit",
            "DocxMarkerControlOrdinarySaveAtomicPublish",
            "DocxMarkerControlOrdinarySaveCountingPublish",
            "PptxMarkerOrdinarySaveLifecycle",
            "PptxMarkerOrdinarySaveEdit",
            "PptxMarkerOrdinarySaveAtomicPublish",
            "PptxMarkerOrdinarySaveCountingPublish",
            "PptxMarkerControlOrdinarySaveLifecycle",
            "PptxMarkerControlOrdinarySaveEdit",
            "PptxMarkerControlOrdinarySaveAtomicPublish",
            "PptxMarkerControlOrdinarySaveCountingPublish",
            "DocxSemanticTextToSink",
            "DocxSourceTextToSink",
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
