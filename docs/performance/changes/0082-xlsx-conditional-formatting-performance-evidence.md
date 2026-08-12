# Change 0082: XLSX conditional-formatting performance evidence

Date: 2026-08-12

Production capability: `b00a34c8670cf6de3ec52a595c4264e860ebeaed`

Status: selectable deterministic evidence; balanced ABBA pending

## Scope and matched operation

The existing `litchi_xlsx::conditional_formatting::SourceBackedEditor` already
provides guarded complete-collection CRUD for direct core conditional
formatting in one existing normal worksheet. This change adds harness and CI
evidence only; it does not change the production editor or its dependency
closure.

Both controls consume the same immutable positional XLSX source and replace
the source's complete ordered collection of three expression-rule owners with
the same three typed target owners. Both use
`replace_conditional_formattings` as their worksheet rewriter. The eager path
first materializes the complete OPC package; the source-backed path loads the
existing guarded workbook, selected worksheet, relationship and styles/DXF
closure and publishes the selected worksheet overlay. Source construction,
sink reservation, typed reopen, hashing and every correctness oracle remain
outside the timed interval.

The source-backed commit is checked for a changed, nonempty patch. A separate
untimed gate applies that exact patch to the source OPC graph, compares the
complete target collection, applies its inverse, and checks that every Part's
content type, relationships and payload are restored exactly.

## Deterministic corpus and oracle

The fixed media-rich workbook has one normal worksheet, one styles Part, one
drawing and eight referenced deterministic incompressible 2 MiB PNG Parts. It
contains 12 logical Parts, 17 ZIP members and 16,783,105 logical Part bytes.
The 16,786,988-byte archive has SHA-256
`843ec3a9fdf759f3ed79064050125975150b16b8230aab52e6074dc000efedfa`.

Both controls produce the same 16,786,983-byte output with SHA-256
`5a44bffa0f3b4ea554e3eaa6588eec770d72528a9c854019c7abc6ee419b9b70`.
The complete untimed oracle reopens the XLSX through the public package and
workbook APIs, checks the three target owners and retained calculation ID,
then compares package topology, package and Part relationships, content types,
every unselected logical Part payload and all eight media payloads. It also
compares each unselected member's exact raw local span and central record
apart from the necessarily relocated local-header offset.

The non-seek output sink reserves its complete budget before timing, caps each
write at 64 KiB and caps total output at twice the expected output plus 64 KiB.
The instrumented positional source requires ordinary payload reads and caps
total source reads at the archive size plus 64 KiB. The deterministic smoke
records 630 eager and 547 source-backed writes, with a 32 KiB largest write.
The eager control materializes all 12 logical Parts; the source-backed control
materializes exactly workbook, selected worksheet and styles (three Parts).

## Reproduction

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 10 --samples 100 \
  --case xlsx_eager_conditional_formatting_edit_save,xlsx_source_backed_conditional_formatting_edit_save \
  --json target/perf/xlsx-conditional-formatting-edit.json
```

The CI smoke uses zero warmups and one sample per control, and asserts the
corpus/output hashes, logical Part and ZIP-member counts, logical/output bytes,
source-read and sink bounds, identical output, and 12-versus-three
materializations.

No latency, instruction, allocation, peak-heap or RSS improvement is claimed.
Those require a retained balanced ABBA protocol and separate counter/memory
evidence.

## Verification

The harness regression constructs the corpus twice, runs one sample for each
control, and checks byte-identical output and exact materialization counts.
The already committed production integration suite separately covers
add/replace/clear/no-op, Strict and Transitional namespaces, complete reopen,
raw unselected-member preservation, exact patch/inverse, signatures,
protection, MCE/x14 refusal, stale and foreign sources, limits and partial
sinks.

Focused commands:

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  xlsx_conditional_formatting_edit_controls_are_deterministic_and_equivalent
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
```

General cells, formulas, x14/MCE-selected conditional formatting, tables,
topology changes, real-producer performance, encryption, atomic filesystem
publication and conditional-formatting clear timing remain outside this case.
