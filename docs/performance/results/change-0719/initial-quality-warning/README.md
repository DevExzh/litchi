# Initial quality warning

This source passes four correctness captures and all 535 tests (one ignored),
but warning-denied Clippy rejects an unnecessary `mut` on the test rewrite
closure argument. The final source removes that binding qualifier and corrects
historical corpus comments. Final builds, captures and quality gates are rerun;
these earlier successful captures are not pooled into final results.
