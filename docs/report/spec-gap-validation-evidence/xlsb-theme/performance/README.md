# XLSB Theme profile

This directory owns the bounded profiling harness for the XLSB Theme owner
(MS-XLSB §2.1.7.52). The Rust example
theme_profile.rs measures one case in one fresh process and emits raw JSON.
Its source lane opens SourceBackedWorkbook through a counted ReadAt, calls
SourceBackedWorkbook::theme(), and retains the returned theme::View for warm
metadata queries. Its eager lane uses Workbook::new and the eager
Workbook::theme() owner. The transaction lanes call Workbook::edit_theme,
ThemeTransaction::commit, and Workbook::apply_theme for the forward commit;
the inverse is published directly with
Workbook::apply_theme_patch(commit.patch().inverse()).

The harness records operation-scoped std::alloc::System requested allocation
and logical ReadAt counters. allocation.peak_after is a requested-live
interval highwater reset to the live baseline before each timed operation. It
is neither RSS nor an allocator-internal highwater. Read counters are source
reader observations, not physical I/O. XML/OPC package loading still follows
the normal package path; these reports do not claim partial ZIP decompression.
Allocator-instrumented elapsed time describes each declared API scope and is
not production latency.

The declared cases are:

- source_cold: source-backed workbook open, selected Theme read, and
  source-view construction. The post-read semantic digest is computed outside
  the timed interval.
- eager_cold: eager workbook open, selected Theme read, and typed snapshot
  construction. The post-read semantic digest is also outside the timed
  interval.
- source_warm and eager_warm: repeated metadata queries against one retained
  typed view/snapshot. Setup is outside the timed samples and is reported with
  its own allocation and source-read counters. Source warm samples retain the
  counted ReadAt state and report cumulative reads before each query plus the
  operation interval; eager warm has no source reader.
- noop: typed no-op transaction commit with exact source-byte preservation.
- change_inverse: typed Accent1 change, readback, inverse publication, and
  exact source-byte restoration. The report is rejected unless semantic change,
  typed readback, inverse state, and source bytes all agree.

No equivalent-work speedup is inferred across source-backed, eager, and
transaction cases because their API validation and lifecycle scopes differ.
The useful evidence is the declared work, allocation/read counters, and
preservation gates for each lane.

The matrix uses the two native fixtures already validated by
verify-theme-schema.py:

    test-data/poi/test-data/spreadsheet/testVarious.xlsb
    test-data/ooxml/xlsb/62815.xlsb

make-theme-control.py creates a deterministic third fixture from
testVarious.xlsb by inserting a 1 MiB XML comment outside the modeled theme
elements. The typed model remains small while raw Theme-part retention and
transaction source-byte work see the larger payload. The runner validates the
generated workbook against the vendored ECMA Transitional schema before any
profile process starts. The comment is not a model-many-fonts workload.

For a compile-only smoke check after the source/test freeze:

    RUSTUP_HOME=/tmp/litchi-spec-gap-rustup CARGO_INCREMENTAL=0 \
      cargo +1.95.0 check --offline --locked -p litchi-xlsb --example theme_profile

The bounded runner always rebuilds when PROFILE_BIN is unset. It captures the
relevant source hashes, Cargo.lock, git HEAD, build log, and binary hash before
and after the matrix. With an externally frozen binary, set both PROFILE_BIN
and PROFILE_BUILD_MANIFEST; the manifest hash is checked before and after the
run.

The final matrix is three fixtures × six cases × three fresh
processes, with three warmups and 30 retained samples per process:

    export RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
    export RUST_TOOLCHAIN=1.95.0 CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=2
    export CARGO_TARGET_DIR=/tmp/litchi-xlsb-theme-target-20260910
    export OUTPUT_DIR="$PWD/docs/report/spec-gap-validation-evidence/xlsb-theme/performance/raw/final-release"
    export PROCESSES=3 WARMUP=3 SAMPLES=30 THEME_CONTROL_BYTES=1048576
    ./docs/report/spec-gap-validation-evidence/xlsb-theme/performance/run-profile.sh

The runner's verifier recomputes
nearest-rank p50/p95/p99 and the mean from raw samples; requires every
semantic/preservation/inverse/change/read-observation gate to be exactly true;
checks allocator and source-counter intervals; and requires the semantic/source
observations to match across the three fresh processes.

The runner records generated-fixture schema output, raw JSON, and provenance
under the selected output directory; it writes matrix-summary.json beside
that directory (under raw/ for the final-release command above). It does not
retain large binaries or claim cache/RSS values that the harness does not
measure.

The verifier was also smoke-tested with synthetic JSON for 3 fixtures × 6
cases × 3 process reports (54 reports). That smoke test exercised only schema,
gate, counter, and percentile validation; its synthetic reports are not
performance evidence and are not retained here.
