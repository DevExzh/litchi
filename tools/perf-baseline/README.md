# OPC, CFB, legacy-writer, and XLSX performance baseline

`litchi-perf-baseline` is an isolated, reproducible measurement tool for the
ZIP/OPC and CFB/OLE2 substrates, fresh DOC/XLS/PPT writer packaging, and
public-API XLSX snapshot/edit/save flows. It creates every corpus in memory; it
also exercises source-backed XLSX catalog and worksheet reads over positional
I/O. It does not depend on untracked office files, network state, randomness,
or the system clock. The JSON report contains the generator parameters and SHA-256
hashes for the generated container and target entry, so a result always
identifies its exact input or packaged output.

The tool is intentionally outside the root workspace and has no effect on
production dependency graphs.

## Run

Run the complete default matrix (36 selectable cases; 198 result records: 144
substrate records, nine writer records, and 45 XLSX records):

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --json target/perf/container-baseline.json
```

For a short local smoke run:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 2 --shape tiny --payload compressible \
  --writer-shape tiny --json -
```

Select comma-separated subsets when investigating a change:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --samples 30 --case cfb_open,cfb_list_streams,cfb_read_one --shape wide-root \
  --payload incompressible --json target/perf/wide-root.json
```

For just the end-to-end legacy writer packaging runs:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to \
  --writer-shape tiny,large,payload-heavy --json target/perf/legacy-writers.json
```

For the dense-wide XLSX range traversal that isolates selected-row scanning:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --case xlsx_narrow_column_range_scan \
  --xlsx-shape dense-wide --json target/perf/xlsx-narrow-range.json
```

For a short positional XLSX laziness and selected-range smoke run:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 \
  --case xlsx_source_open,xlsx_source_list_sheets,xlsx_source_first_cell,xlsx_source_narrow_column_range_scan \
  --xlsx-shape tiny --json -
```

## Corpus matrix

Each shape is generated twice, once with a repeated deterministic payload and
once with a deterministic xorshift payload. The latter is intended to be hard
to compress, not cryptographically random.

| Shape | Entries | Bytes per entry | Total logical payload | Primary CFB stressor |
|---|---:|---:|---:|---|
| `tiny` | 3 | 512 B | 1.5 KiB | Fast smoke input; CFB MiniFAT |
| `many-small` | 256 | 1 KiB | 256 KiB | CFB MiniFAT traversal |
| `few-large` | 4 | 4 MiB | 16 MiB | CFB regular FAT streams |
| `wide-root` | 2,048 | 64 B | 128 KiB | Wide CFB root sibling tree |

## Legacy writer matrix

The fresh writer cases are independent of the ZIP/CFB corpus shapes and
payload kinds. `--writer-shape` selects their deterministic bounded content
shape; omitting it runs all three.

| Writer shape | DOC paragraphs | XLS cells | PPT slides / text boxes | Bound |
|---|---:|---:|---:|---:|
| `tiny` | 3 | 16 | 1 / 2 | Fast public-API smoke input |
| `large` | 512 | 8,192 | 12 / 144 | Moderate end-to-end packaging input |
| `payload-heavy` | 128 × 20,000-byte text | 128 × 32,700-byte string cells across 128 sheets | 16 / 128 × 40,000-byte text | 4–8 MiB primary CFB stream |

DOC paragraphs and PPT text boxes contain fixed identifier text. XLS uses a
fixed sequence of finite numeric cells for `tiny`/`large`. `payload-heavy`
uses repeated deterministic text; its XLS string cells are distinct but each
worksheet contains exactly one cell, retaining stable shared-string ordering
through the public API. There is no random, clock, or filesystem-derived
content.

## XLSX corpus matrix

`--xlsx-shape` selects complete, in-memory workbooks with a fixed integer grid
on every sheet. The value at each coordinate is derived from its sheet, row,
and column. The deterministic ~1% update set is evenly distributed through the
complete logical cell order and rounds up to one cell where necessary.

| XLSX shape | Sheets | Rows × columns per sheet | Logical cells | ~1% updates | Primary use |
|---|---:|---:|---:|---:|---|
| `tiny` | 3 | 8 × 8 | 192 | 2 | Fast end-to-end smoke input |
| `medium` | 4 | 32 × 32 | 4,096 | 41 | Moderate snapshot/edit/save input |
| `dense-wide` | 2 | 256 × 256 | 131,072 | 1,311 | Stable narrow-column scan over dense selected rows |

`dense-wide` deliberately contains all 256 stored cells in each selected row.
Its narrow range is `B1:B256`, which returns 256 cells while making the
worksheet store examine the 65,536 stored cells in those rows. This preserves a
high-contrast public end-to-end case without relying on internal APIs.

OPC parts retain their deterministic `benchmark/parts/00000.bin` names, and
the middle entry remains the fixed `zip_read_one` target. CFB streams are
fixed-width root siblings named `benchmark_stream_00000.bin`, etc. The final,
lexicographically greatest CFB stream is the `cfb_read_one` target. This makes
`wide-root` exercise a long successful validated-tree descent instead of
accidentally measuring a near-root hit, and keeps the original pre-change
full-tree traversal cost reproducible in comparison binaries.

Every synthetic OPC corpus has exactly one package-level Office Document
relationship, `rIdBenchmarkMain`, targeting that middle Part. These packages
use generator identifier `litchi-opc-synthetic-v2`; older v1 corpus hashes
remain distinguishable.

## Cases

- `zip_index`: parse the ZIP central directory and build Soapberry's index.
- `zip_read_one`: read and verify one target member from an already indexed ZIP.
- `opc_open`: parse a generated OPC ZIP into `litchi_opc::OpcPackage`.
- `opc_open_owned`: move an already-prepared archive allocation into
  `OpcPackage`; the input clone happens before timing.
- `opc_noop_save`: open the owned archive bytes once, then serialize the
  unchanged package through a bounded non-seekable counting sink. This measures
  Litchi's exact owned-source no-op path and verifies byte-identical output.
- `opc_mutated_save`: change one byte in the fixed target Part before timing,
  then publish through the full validated rewrite path and compare every output
  byte with a deterministic expected archive.
- `opc_source_open`: build the positional source's validated catalog with an
  explicit finite cache policy. A post-timing first payload access must still
  perform source I/O, proving that open did not materialize it.
- `opc_source_open_main_read`: open the same source and materialize its unique
  Office Document main Part in one timed operation.
- `opc_source_cached_main_read`: cold-load before timing, then require the timed
  access to add zero source reads and return the same pinned `Arc` allocation.
- `opc_source_concurrent_same_part`: start two cold main-Part reads together;
  both must receive the same pinned allocation, and a third access must be a
  zero-I/O cache hit.
- `cfb_open`: parse the complete generated container into `litchi_cfb::OleFile`.
- `cfb_list_streams`: enumerate and materialize all stream paths from an
  already-open CFB container.
- `cfb_read_one`: look up, materialize, and verify the final root stream from an
  already-open CFB container.
- `cfb_create_stream_borrowed`: insert one already-prepared target payload with
  `OleWriter::create_stream`, including its required payload copy in the timed
  interval.
- `cfb_create_stream_owned`: insert the same already-prepared target allocation
  with `OleWriter::create_stream_owned`. Comparing this directly with the
  borrowed case isolates the avoidable payload-copy cost.
- `cfb_shared_open`: validate and index a CFB container through
  `SharedOleFile` over the instrumented positional source.
- `cfb_shared_read_one`: open outside timing, reset counters, then read and
  verify the fixed target stream through positional I/O.
- `cfb_shared_concurrent_reads`: start reads of the first and final root streams
  together after open. The result records observed maximum in-flight reads
  without requiring overlap on very small corpora.
- `doc_fresh_write_to`: construct a new `litchi_doc::writer::Writer`, add the
  selected fixed paragraphs through its public API, and package it with public
  `write_to`.
- `xls_fresh_write_to`: construct a new `litchi_xls::writer::Writer`, add the
  selected sheets and cells through its public API, and package it with public
  `write_to`.
- `ppt_fresh_write_to`: construct a new `litchi_ppt::writer::Writer`, add the
  selected slides and text boxes through its public API, and package it with
  public `write_to`.
- `xlsx_open_owned`: move a prepared owned XLSX allocation into public
  `Workbook::from_bytes`; the input clone happens before timing.
- `xlsx_list_sheets`: enumerate sheet names after an already-open workbook,
  without asking any worksheet for cell semantics.
- `xlsx_first_cell`: access the first stored cell (`Sheet1!A1`) from an
  already-open workbook, measuring the first worksheet semantic access.
- `xlsx_full_cell_scan`: enumerate every stored cell in `Sheet1` through the
  public full-sheet range.
- `xlsx_narrow_column_range_scan`: preload `Sheet1` outside timing, then scan
  `B1:B<rows>`. On `dense-wide`, that returns 256 cells while traversing all
  65,536 stored cells in the selected rows.
- `xlsx_noop_commit`: commit a fresh public edit with no mutations and verify
  its patch is empty and its complete workbook semantics are unchanged.
- `xlsx_noop_commit_save`: do the no-op commit and public `Workbook::write_to`
  into the bounded in-memory sink, then require exact generated-corpus bytes.
- `xlsx_one_cell_commit` / `xlsx_one_cell_commit_save`: prepare one fixed cell
  update before timing, then time commit alone or commit plus public
  `write_to`; the resulting patch, workbook semantics, and save bytes are
  verified.
- `xlsx_one_percent_commit` / `xlsx_one_percent_commit_save`: do the same for
  the shape's deterministic ~1% update set (2 / 41 / 1,311 cells).
- `xlsx_source_open`: open `SourceBackedWorkbook` over the instrumented
  positional source. A fresh post-timing `Sheet1!A1` read must add selected
  worksheet member I/O, proving that open did not semantically materialize a
  worksheet.
- `xlsx_source_list_sheets`: enumerate the source-backed catalog after open,
  require zero timed source reads, and use the same fresh first-cell delta to
  prove that listing did not materialize a worksheet.
- `xlsx_source_first_cell`: cold-read and verify `Sheet1!A1`, then require a
  fresh second-sheet access to add unselected worksheet member I/O. This proves
  that the selected read did not semantically materialize another worksheet.
- `xlsx_source_narrow_column_range_scan`: cold-read and verify every address
  and value in `B1:B<rows>` through the source-backed public API, with the same
  second-sheet deferral proof.

For both CFB stream-insertion cases, payload generation/cloning and writer
construction happen before timing, while writer and source destruction happen
afterward. They time stream insertion only, not complete CFB serialization. The
`few-large` shape is the useful ownership profile because its prepared target is
4 MiB; `tiny` stays small enough for a quick release-build smoke test.

The sink reserves a checked byte budget before timing and copies every output
byte into bounded memory. It records accepted bytes, write count, and the
largest write, rejects output beyond the budget or individual writes larger
than 64 KiB, and verifies the deterministic no-op output after timing. This
keeps allocator growth out of the timed interval while preventing a
passthrough write from degenerating into bookkeeping that never reads bytes.

Each fresh writer iteration begins with a new writer and ends after its public
`write_to` call has completed into a seekable `Cursor<Vec<u8>>`. This exercises
the real DOC/XLS/PPT OLE packaging path used by the corresponding `save` API
without timing filesystem state. Construction and packaging are both included;
post-timing validation requires the package bytes to exactly match the
precomputed deterministic corpus. The manifest records the complete output
hash and the hash of its conventional primary stream (`WordDocument`,
`Workbook`, or `PowerPoint Document`).

Each XLSX save case uses public `Workbook::write_to` rather than filesystem
`save`, so it measures the same sequential serialization path without timing
filesystem state. No-op byte exactness is deliberately scoped to the harness's
generated XLSX corpus, whose owned-source serialization is deterministic; it
does not claim byte preservation for arbitrary third-party XLSX input.

## JSON contract

Reports use `schema_version: 1`. Each result contains sorted raw elapsed-ns
samples plus min, median p50, nearest-rank p95 and p99, max, mean, sample
standard deviation, and a two-sided Student's-t 95% confidence interval for
the mean. Untimed warm-up iterations are recorded in the configuration but are
never included in these statistics.

Each case embeds a corpus manifest with the format-specific generator version,
container format, shape, payload algorithm, logical and serialized sizes,
member count, target member, and SHA-256 hashes. Existing OPC manifest fields
are unchanged; CFB reports use `compression: "none"`
and count streams in `archive_member_count`. The report also records the git
revision and dirty state when available, `rustc` version, build profile and
target, visible logical CPUs, allocator, relevant Cargo/Rust flags, and Linux
`perf_event_paranoid` value. Metadata collection is best-effort and is complete
before timed iterations. The requested JSON parent directory is created
automatically.

`configuration.writer_shapes` is an additive schema-v1 field identifying the
fresh writer shape selection. For writer records, `entry_count` is the logical
paragraph/cell/text-box count, `uncompressed_payload_bytes` is the generated
logical content byte count, and `entry_bytes` is zero because the content has
no uniform per-entry byte size. Existing substrate fields and generator
identifiers retain their previous meanings.

`configuration.xlsx_shapes` and the optional `corpus.xlsx` object are additive
schema-v1 fields. `corpus.xlsx` records sheet dimensions, the exact
deterministic ~1% update count, and `source_members`: the workbook, worksheet,
shared-string, and style ZIP member names whose exact compressed ranges feed
the positional counters. Existing non-XLSX corpus records keep it null.

The eleven positional cases add a `source` object; older cases omit it. Its
arrays contain one value for every measured iteration and record `read_calls`,
`read_bytes`, compressed ordinary-OPC-payload range overlap, and
`max_in_flight_reads`. OPC cases also record semantic
`ordinary_payload_materializations` (0 or 1). Payload-range overlap is a
physical request-amplification metric: bounded ZIP metadata reads may fetch
adjacent compressed payload bytes without decompressing or caching that Part.
Accordingly, `opc_source_open` may report overlap while still reporting zero
materializations; its post-timing cold access proves the distinction.

Positional XLSX records additionally contain `source.xlsx` arrays for physical
overlap with the workbook, selected worksheet, all unselected worksheets,
shared strings, and styles compressed member ranges. These overlap counters
are intentionally truthful about ZIP read amplification and therefore are not
semantic materialization counters. No XLSX materialization count is emitted,
because the production API does not directly expose one. Instead, each case
enforces its semantic deferral claim with a fresh post-timing worksheet access
that must add I/O for that worksheet's exact compressed member range.

## External profiling

Build the binary once before collecting counters, then invoke it directly so
Cargo is not part of the profile:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
perf stat -d tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --case cfb_read_one --shape wide-root \
  --payload incompressible --json target/perf/perf-cfb-read.json
```

For a sampled profile, use whichever installed tool is appropriate:

```sh
cargo flamegraph --release --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 --case opc_noop_save --shape few-large \
  --payload compressible --json target/perf/flamegraph-run.json

samply record tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 3 --samples 15 --case opc_noop_save --shape few-large \
  --payload compressible --json target/perf/samply-run.json
```

External profilers are optional. On Linux, `perf` can be installed yet denied
by the kernel's `perf_event_paranoid` policy or container permissions. The
harness records that policy when readable and still runs normally without any
profiler, providing the wall-clock JSON baseline. If `perf` is unavailable or
denied, run the binary directly (or use `samply`/`cargo flamegraph` if present)
rather than changing system policy merely to complete a smoke run.

Timing is only a first baseline. It intentionally does not claim peak RSS,
allocation counts, CPU utilization, lock contention, cache misses, or scaling;
those need dedicated ADR-0005 instrumentation and a controlled runner.
