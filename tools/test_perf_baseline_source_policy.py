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
        cls.metrics = (
            PERF_BASELINE / "src" / "allocation_metrics.rs"
        ).read_text(encoding="utf-8")
        cls.filesystem = (PERF_BASELINE / "src" / "filesystem.rs").read_text(
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
        self.assertIn("use litchi_perf_baseline::allocation_metrics;", self.allocator)
        self.assertNotIn("use super::allocation_metrics;", self.allocator)
        self.assertIn("unsafe impl GlobalAlloc", self.allocator)
        self.assertIn("#[global_allocator]", self.allocator)
        self.assertIn("System.alloc", self.allocator)
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
