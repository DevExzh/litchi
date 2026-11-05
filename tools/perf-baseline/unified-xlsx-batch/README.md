# Paired XLSX scalar-cell batches

This isolated harness measures one native XLSX transaction against the
unified `BatchItem` route for distinct scalar cell assignments. It builds a
blank workbook with `litchi_xlsx::Workbook::create()`, prepares all addresses
and command payloads before timing, and uses counts 1, 10, 100, 1,000 and
10,000. The native leg performs one `Edit::sheet("Sheet1")` selection and one
`WorksheetEdit::set` loop followed by one commit. The unified leg submits the
same values as an atomic `BatchItem` set batch.

Build and run it from the repository root:

```sh
CARGO_NET_OFFLINE=true cargo build --release --locked --manifest-path tools/perf-baseline/unified-xlsx-batch/Cargo.toml
taskset -c 0 tools/perf-baseline/unified-xlsx-batch/target/release/unified-xlsx-batch-baseline 9 > /tmp/unified-xlsx-batch.json
```

The fixed route names are `xlsx_eager_leaf_scalar_batch_v1`,
`xlsx_eager_unified_scalar_batch_v1`, `xlsx_source_leaf_scalar_batch_v1`, and
`xlsx_source_unified_scalar_batch_v1`. The eager pair opens identical prepared
owners. The source pair opens the same bytes through the retained source
owner; the native source owner is promoted to an exact-source `Workbook`
before timing, and the first unified source batch is run as an untimed warm-up
so source promotion is complete before ABBA samples begin.

Each count is warmed once, then measured for nine ABBA rounds by default, with
the order native, unified, unified, native. The runner verifies the native and
unified output bytes, one native patch change and one changed unified result
per item, and one commit per route before timing. Both measured outputs retain
their complete native commit or unified result until the timer has stopped;
workbook opening, source promotion, command construction, serialization, and
output destruction remain outside the interval. JSON includes the corpus size
and SHA-256, route names, sample arrays, medians, operation/item/commit counts,
and the source-promotion warm-up statement.

This is scoped evidence for distinct scalar cell batches. It does not claim
performance for repeated targets, ranges, structural workbook edits, saving,
or other formats.

The record-0101 corpus is 2,077 bytes with SHA-256
`55f12476f7b54eb64b9f946750cb5ce25eee5db4393e8f43c0205d25202a64bc`.
Changing the corpus or route semantics requires new route names.

`without-promoted-resolution.patch` reconstructs the measured control: native
commands still coalesce, but each command resolves through the retained source
instead of the already-promoted native owner. Apply only to matching source,
build and measure, reverse it, then rebuild the candidate:

```sh
git apply --check tools/perf-baseline/unified-xlsx-batch/without-promoted-resolution.patch
git apply tools/perf-baseline/unified-xlsx-batch/without-promoted-resolution.patch
# Run the build and timing commands above; preserve results outside the checkout.
git apply -R tools/perf-baseline/unified-xlsx-batch/without-promoted-resolution.patch
# Rebuild before measuring the candidate.
```

Record 0101 retains two nine-round invocations for each variant, compiler and
machine details, lockfile/source hashes and binary identities. Build binaries
are removed after measurement and must never be committed. This is an isolated
control, not a reconstruction of an earlier repository commit.
