# Selected-cell query control protocol (0959)

Status: **prospective and unexecuted**. This note records a control protocol only; no
control build or workload was launched for this review. The active 0959 managed-
compression capture remains separate and untouched.

## Purpose and scope

The control isolates the two existing marker-free XLSX producer selected-cell
routes affected by the Arc-backed `Number` candidate:

- `xlsx_producer_medium_control_selected_cell`
- `xlsx_producer_dense_control_selected_cell`

The cases use the existing `litchi-perf-baseline` runner and generated in-memory
`OwnedSource` corpora. They are warm/in-memory source controls; they make no cold
filesystem claim. The control is diagnostic and has no independent adoption gate.
It must not be used to claim public `Number::new` performance or an allocation
saving.

## Frozen source and binary identities

The source roots are:

- Baseline: `/home/zhuhe/litchi-performance-0959/source-baseline`
- Candidate: `/home/zhuhe/litchi-performance-0959/source-candidate-arc-v1`

The candidate source is the reviewed three-path delta. Its production file hashes
are:

- `crates/litchi-xlsx/src/cell.rs`
  `f345c0ab2101d5dad2b25da4bfbfcb3f1bdafabdeca2ad31f66535887e5c2291`
- `crates/litchi-xlsx/src/raw/worksheet/selected.rs`
  `3da16a84cf9f3cc8b36eb75baea359c17bcf969bb6f78a702452a74f4fa13a63`
- `crates/litchi-xlsx/src/raw/worksheet/semantic.rs`
  `bd357cedab881148680fe389e7d4288c954239cfdae5f41dbba09e1327f97ebd`

The reviewed source manifests are:

- Baseline comparison manifest:
  `/home/zhuhe/litchi-performance-0959/source-manifest-baseline-comparison-v1.json`
  SHA-256 `7397cc28dbaa29e3b6f43603c26cedcd7f92e9bc57b865d3ae6acf0d01418f79`
- Candidate manifest:
  `/home/zhuhe/litchi-performance-0959/source-manifest-candidate-arc-v1.json`
  SHA-256 `77ec8f3efe5c0dfb9aebf6a66ddcd061718c45c532388e7c26297fee9dc508cc`

Both source `Cargo.lock` files are 138443 bytes and have SHA-256
`d80aa535dcf880a85fa31e92fe2e5ce78b91d93a55953ae734e513aed83e0061`.

The current active 0959 capture binaries are managed-compression binaries and
are not valid binaries for this control:

- Baseline `/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2`, 43259328 bytes, SHA-256 `5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`
- Candidate `/home/zhuhe/litchi-performance-0959/executables/xlsx_managed_compression_benchmark.arc-v1`, 43228520 bytes, SHA-256 `3cf1714372b376ccdaeaf6a631a5439195694bdfffd4d10c25dc5f06db95f51f`

A future control must build and independently bind these separate targets:

- `/home/zhuhe/litchi-performance-0959/target-query-baseline/release/litchi-perf-baseline`
- `/home/zhuhe/litchi-performance-0959/target-query-candidate/release/litchi-perf-baseline`

The planned build command, run from each corresponding source root, is:

```text
cargo +1.95.0 build --release --locked --offline \
  --manifest-path <source>/tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline
```

The build environment must match the frozen 0959 profile: `RUSTUP_TOOLCHAIN=1.95.0`,
`CARGO_BUILD_JOBS=2`, `CARGO_INCREMENTAL=0`,
`CARGO_PROFILE_RELEASE_OPT_LEVEL=3`, `CARGO_PROFILE_RELEASE_LTO=fat`,
`CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1`,
`CARGO_PROFILE_RELEASE_DEBUG=1`, `CARGO_NET_OFFLINE=true`, `LANG=C`, `LC_ALL=C`,
and `TZ=UTC`. Use separate `CARGO_TARGET_DIR` and temporary directories for the
two arms. No build is authorized by this note.

## Paired execution

Use six paired blocks if the control is promoted beyond a pilot; three paired
blocks are a minimum descriptive pilot. Each block launches two fresh processes,
one baseline and one candidate, alternating `A,B,A,B,A,B` order across blocks.
Use the same fixed CPU mask as the active 0959 protocol, CPUs `12-15`, and never
run two children concurrently. Each child runs both cases in the same explicit
case order and uses 30 warmups and 200 retained samples:

```text
/usr/bin/taskset --cpu-list 12-15 <binary> \
  --case xlsx_producer_medium_control_selected_cell,xlsx_producer_dense_control_selected_cell \
  --warmup 30 --samples 200 \
  --json <unique-report.json> \
  --producer-evidence <unique-producer-evidence.json>
```

The output directory must be external to the repository, with unique absent paths
for every child. Retain only the two report rows, the producer-evidence sidecar,
terminal/exit custody, command, cwd, environment subset, binary descriptor, and
source descriptor for this scoped control. Do not run the allocator binary or the
broader default matrix.

## Report and semantic oracles

The runner report is schema version 1 (`Report` in `tools/perf-baseline/src/lib.rs`)
and must satisfy all of the following:

1. `configuration.cases` contains exactly the two requested case names in that
   order, with 30 warmups and 200 samples.
2. `results` contains exactly two rows in the same order. Each row has
   `elapsed_ns.unit == "ns"`, 200 samples, and a `sample_order` permutation of
   `0..200`. The statistics samples are elapsed-sorted; they are not chronological
   sample order. No `source`, `sink`, `execution`, `output_sha256`, or
   `operation_metrics` field is expected for these `plain_result` rows.
3. The complete `corpus` manifest for each case is byte-for-byte/equivalent JSON
   across baseline, candidate, and blocks, including archive size/SHA, target
   payload size/SHA, XLSX dimensions, and source-member names.
4. The sidecar has schema
   `litchi.perf-baseline.producer-shape-evidence.v1` and exactly two corpora,
   medium-control and dense-control. Its full JSON must be equal across arms and
   blocks. The selected target/address and `selected_target_text` must match; the
   selected worksheet must report `source_stream_eligible: true`, no ineligible
   reasons, and no `<cols>` exclusion.
5. Successful process exit is a semantic check: every measured iteration calls
   `describe_cell(cell.stored()) == expected`. A mismatch is a child failure, not
   a result to be averaged. No replacement or retry is permitted.
6. Binary hashes are expected to differ. Compare the bound source manifests,
   lock hash, Rust/profile settings, CPU affinity, command, and tool schema rather
   than requiring executable identity equality.

For analysis, pair the baseline and candidate distributions by block and case;
report raw values and p50/p95 summaries with their existing confidence intervals.
Do not interpret this control as a public API or allocation measurement, and do
not create a new acceptance threshold after observing the data.

## `Number` representation risk

The candidate changes `Number` from `Box<str>` to `Arc<str>`. Public
`Number::new(value: impl Into<Box<str>>)` remains unchanged. Candidate
`Number::new` first materializes the `Box<str>` and then copies into an `Arc<str>`
(`crates/litchi-xlsx/src/cell.rs:398-402`), so external callers can still incur a
second allocation/copy. The private `Number::from_lexical(&str)` path allocates
one Arc directly.

The selected scanner is not a pure public-constructor test. Baseline
`selected.rs` constructs the selected numeric value with
`Number::new(value.into_boxed_str())`; candidate constructs it with
`Number::from_lexical(&value)`. Since `value` is already an owned scratch
`String`, candidate copies into a new Arc while baseline can transfer the string
allocation into a Box. A neutral or negative selected-cell result is therefore
possible and would be useful route evidence, not proof that the Arc design is
invalid elsewhere.

No existing timed benchmark was found for the public constructor. The only
benchmark-harness call is `tools/perf-baseline/src/xlsx_planning_guard.rs:502`,
outside its timed `edit_sheets` interval. The `perf_cap_boundary` example creates
expected `Number::new` values after its timer (`crates/litchi-xlsx/examples/perf_cap_boundary.rs:371-372`). The managed-compression benchmark sources contain no public `Number::new` call. Consequently, adoption must make no public-constructor performance claim based on this control.
