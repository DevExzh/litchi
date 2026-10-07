# 0958 XLSX numeric ownership trace

## Scope and evidence boundary

This is a read-only source trace at `e29699f49af54e435145cb6051668f6929e7e16f`.
It covers the source-backed XLSX cell-value editor and the deterministic medium
and dense-sparse scalar-grid corpus. iWork is outside the scope. No build,
benchmark, or source edit was performed for this note.

The allocator evidence is the retained 0957 report
`/home/zhuhe/litchi-performance-0957/allocation-summary-v3.json`, SHA-256
`bc311f3209f346baacedd591a17d6fb36937dda81a256e1ea1cb4f2aaabd4850`. Its
contract is system-allocator operation regions and observer-ordered live peaks;
it does not attribute a byte to a source call stack. The current source file
SHA-256 values are recorded below so a later heaptrack or allocator attribution
can bind to this exact tree.

The source-derived counts below are exact for the admitted provenance route and
this generated corpus. They are not a claim that the whole phase counter belongs
to `Number`, and the lexical-byte sums are not allocator counter readings.

## Finding

`Store::merge_omitted_cells` is the direct numeric clone site.

* `Number` is a public opaque wrapper around `Box<str>` and derives `Clone`
  (`crates/litchi-xlsx/src/cell.rs:388-397`). `Value`, `Cell`, and `Stored` also
  derive `Clone` (`cell.rs:192-202`, `cell.rs:82-95`, `cell.rs:646-676`).
* The source-backed value-only rewrite records unchanged worksheet runs as
  `OmittedCells`. Untouched rows and the unchanged runs surrounding a changed
  cell are recorded by `write_sheet_data_with_provenance`
  (`crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs:151-195,
  384-453`). This is independent of replacement-span hints; the omission list is
  used for reduced semantic readback.
* `Snapshot::try_rewritten_value_cells` parses the reduced candidate, converts
  each omission to a checked rectangle, and calls `Store::merge_omitted_cells`
  (`crates/litchi-xlsx/src/cell_values/snapshot.rs:832-859`). The merge moves the
  reduced parsed cells, then executes `cells.push(entry.clone())` for every
  omitted source entry (`crates/litchi-xlsx/src/cell.rs:893-959`). On this
  corpus every stored cell is a numeric `Cell::Value(Value::Number(_))`, so each
  such `Stored::clone` clones its `Box<str>`.
* The resulting vector is passed to `Store::from_unsorted`, which sorts and
  rebuilds indexes (`cell.rs:754-818`, called at `cell.rs:936-943`). The sort and
  index work are separate costs; they do not create another numeric lexical
  allocation.

The rest of the snapshot and commit ownership path does not add another
`Number` box for these records. `Snapshot::from_rewritten_value_source` first
clones the source snapshot, but replaces its `cells` with the newly parsed
`Arc<Store>` (`snapshot.rs:795-829`); the source snapshot's existing store is
shared by `Arc`, so that snapshot clone does not clone its cells. The multi-edit
then freezes the candidate, makes the patch, and stores the same target state in
both the commit and the patch (`crates/litchi-xlsx/src/cell_values/source.rs:1146-1254`,
`crates/litchi-xlsx/src/cell_values/patch.rs:217-270`).
`MultiSnapshot::clone` at `source.rs:1248` clones `Snapshot` metadata and the
`Arc<Store>`; it does not clone the `Number` boxes already in that store. The
publication route has another target snapshot clone at `source.rs:639-650`, with
the same property. The candidate numeric boxes are finally released when the
candidate store's last owner is dropped, which is why this work also appears in
the commit-drop deallocation region.

## Deterministic corpus census

The generator builds four worksheets. Medium uses a 48 by 48 full grid per
worksheet (`tools/perf-baseline/src/lib.rs:22660-22710`,
`22879-22937`), and the cell lexical is
`sheet * 1_000_000 + row * 1_000 + column`
(`lib.rs:24295-24300`). The one-percent count is the ceiling of total cells
 divided by 100 (`lib.rs:22713-22733`). All updates are existing-cell updates;
none is an insertion.

| case | source numeric owners | updates | changed by sheet | omitted source numeric owners | omitted lexical bytes |
| --- | ---: | ---: | --- | ---: | ---: |
| medium, one edit | 2,304 in the selected sheet | 1 | 1 | 2,303 | 10,933 |
| medium, one percent | 9,216 | 93 | 24 / 23 / 23 / 23 | 9,123 | 58,723 |
| dense-sparse, one percent | 17,792 | 178 | 164 / 11 / 2 / 1 | 17,614 | 92,907 |

For the medium one-percent row, all four sheets are touched. Each changed row
records every unchanged cell around its update, and every untouched row records
its full cell run. The row-level omission reservation is bounded by the source
row count and succeeds for this corpus, so the omitted count is total owners
minus updates. The 58,723-byte sum is the exact sum of the original numeric
lexeme lengths for those 9,123 values. A derived call-stack report should find
one `Box<str>` allocation per omitted numeric clone and a requested payload size
matching this sum, subject to allocator-layout accounting and any wrapper
frames.

The medium one-edit path has the same ownership pattern on one sheet: 2,303
numeric clones and 10,933 lexical bytes. Dense-sparse gives the larger
source-derived count, but this task's primary phase evidence is medium.

## Relation to 0957 phase counters

The worker-one medium allocation rows are deterministic across the three
repeats:

| case | phase | allocation calls | allocated bytes | deallocation calls | deallocated bytes |
| --- | --- | ---: | ---: | ---: | ---: |
| medium, one percent | edit_commit | 110,124 | 17,166,680 | 90,875 | 12,903,295 |
| medium, one percent | commit_drop | 0 | 0 | 18,703 | 4,262,881 |
| medium, one edit | edit_commit | 27,327 | 3,949,315 | 22,481 | 2,899,015 |
| medium, one edit | commit_drop | 0 | 0 | 4,725 | 1,049,960 |

The numeric clone subset is therefore 9,123 calls / 58,723 lexical bytes for
medium one-percent, or 2,303 calls / 10,933 bytes for medium one-edit. The
phase totals include parsing, candidate buffers, metadata, index rebuilding,
validation, and other ownership. The report cannot prove the subset's timing
share. Width-two and width-four commit-drop allocation/deallocation differences
are pool/teardown observations and are excluded from this attribution.

The 0957 report also shows that worker-one medium one-percent `commit_drop` has
no allocations and 18,703 deallocations. That is consistent with the candidate
numeric boxes being allocated during `edit_commit` and released later; it does
not isolate them from the other candidate objects released in that region.

## Construction paths

There are three relevant numeric construction paths:

1. The complete semantic parser calls `Number::new(value)` for numeric `<v>`
   lexicals (`crates/litchi-xlsx/src/raw/worksheet/semantic.rs:168-185`). A
   source open therefore owns one boxed lexical per parsed number.
2. The selected scanner validates unselected numerics without retaining a
   `Number`, but materializes selected numerics with
   `Number::new(value.into_boxed_str())`
   (`crates/litchi-xlsx/src/raw/worksheet/selected.rs:672-702`). Its fallback
   parser similarly calls `value.to_owned().into_boxed_str()` before `Number::new`
   (`selected.rs:2117-2125`).
3. The reduced value-only readback parses the changed records, then the
   provenance merge clones unchanged source `Stored` records. The first two
   paths allocate newly parsed lexicals; the third allocates cloned lexicals.
   `Store::from_unsorted` moves these records and does not clone their `Number`
   values.

The medium corpus is integer-only and has no shared-string part. Every numeric
lexeme is short, but the public type must retain arbitrary valid lexical forms
and reject non-finite values. The current `Number` representation therefore
preserves exact spelling, including whitespace, exponent notation, and signed
zero; any replacement must preserve these behaviors.

## Narrow candidate options

### Shared numeric backing (`Arc<str>`)

Changing the private backing of `Number` from `Box<str>` to `Arc<str>` would make
`Stored::clone` share the lexical allocation. `Arc<str>` and `Box<str>` are both
fat-pointer values, so this does not by itself enlarge each `Number` field. It
would preserve `as_str`, equality, validation, exact write spelling, and the
`Send + Sync` snapshot requirement. This is a candidate representation change
that could remove the merge clone allocation without changing `Store` or
`MultiSnapshot` ownership; its net value remains open pending the evidence
described below.

The parser must construct the `Arc` directly from the borrowed lexical. The
current selected path first creates a `Box<str>` (`selected.rs:688-690`), and a
naive `Arc::from(value.into_boxed_str())` would retain the temporary Box
allocation and allocate again for the Arc. A private source-parser constructor
that accepts `&str`, plus analogous changes at `semantic.rs:178` and
`selected.rs:688`/`2123`, avoids that conversion. The public `Number::new`
contract must be reviewed before changing its input bound; preserving an
`Into<Box<str>>` public path while converting to Arc inside it would make
ordinary callers pay the two-allocation conversion. A direct `Arc` constructor
or a carefully source-compatible generic boundary is required.

Arc has a real tradeoff. Each distinct number now carries an Arc control block,
while the generated corpus has almost no repeated numeric lexemes. Source
opening and retained live bytes could rise by the control-block overhead for
roughly 9,216 medium values; the one-percent edit would then save the 9,123
candidate lexical allocations. The 0957 setup/live counters and a type-size
check must decide whether the lower edit/commit allocation count compensates for
that retained overhead. This is a candidate for the next heaptrack/paired
allocation run, not a measured improvement in this note.

### Inline short lexemes

A small inline representation with a boxed fallback could avoid heap allocation
for this corpus's short integer lexicals and make omitted-record clones copy
bytes in place. It is a layout-sensitive change: the current `Number(Box<str>)`
is one fat pointer, while an inline-plus-fallback enum carries an inline buffer,
length, and fallback pointer/discriminant. The enum's maximum variant and
alignment determine whether `Number`, `Cell`, or `Stored` actually grow; that
must be measured with `size_of::<Number>()`, `size_of::<Cell>()`, and
`size_of::<Stored>()` before choosing a capacity. It is not safe to assume that
every `Stored` grows, or that no layout grows. `as_str` must return a stable
borrowed view for both variants, and all exact lexical validation/write paths
need coverage.

The historical survey already lists short-number inlining as a smaller
opportunity and cites the 0539 result as the likely low-impact outcome
(`docs/performance/0587-remaining-opportunity-survey.md:676-684`). The 0539
candidate reduced short-lived attribute-string allocation calls but failed the
workflow gates and was restored (`docs/performance/changes/0539-xlsx-transient-ownership-rejected.md:18-35`).
That is coverage against a broad prior candidate, not proof that current
`merge_omitted_cells` numeric clones are immaterial. The new source-derived
9,123-clone count justifies a fresh attribution check, while the required
before/after type-size and retained-live-byte checks keep the inline option
bounded.

### Store overlay or persistent cell backing

The historical 0556 preparation was a checked **linear merge** candidate for
`Store::merge_omitted_cells`: it replaced append-then-sort while retaining the
ordinary constructor's validation and index rebuild. It did not implement an
`Arc<Store>` overlay (`docs/performance/0556-xlsx-sorted-merge-preparation.md:1-27`).
The later 0587 survey proposed an `Arc<Store>` plus sorted overlay, retained
roughly 5.97% Store-merge attribution and modelled about 16,000 one-edit commit
allocations, but that overlay was not implemented. The subsequent 0557 pilot
stopped on baseline noise before candidate measurement
(`docs/performance/0557-xlsx-allocation-noise.md:1-3,13-32`). An overlay would
remove numeric entry clones and append-then-sort/reindex work, but it is a
broader Store design with more read-path and bound obligations. These historical
records establish prior coverage and limits; they do not select an overlay or a
numeric representation from the 0957 counters.

## Next evidence and decision gate

Use heaptrack or an equivalent call-stack allocator capture on the current
medium one-percent worker-one path and require the following source-derived
witnesses in the report:

* `Store::merge_omitted_cells` / `Stored::clone` / `Number` clone frames;
* approximately 9,123 numeric clone allocations and 58,723 lexical payload
  bytes for the medium one-percent admitted route;
* the separate reduced-parser allocations for the 93 changed values;
* no attribution of width-two/width-four pool teardown to this candidate.

If the stack confirms the expected subset, compare the unchanged Box baseline
with direct-parser `Arc<str>` and inline-plus-fallback candidates under fresh
source/package ownership. Keep the decision open until creation cost, setup and
retained-live bytes, type sizes, and edit/commit allocation deltas are measured.
The public `Number::new` input bound remains `Into<Box<str>>` unless a deliberate
source-compatible API decision is made; a private direct-parser Arc constructor
avoids a parse-side Box-to-Arc conversion, but it does not make the public
constructor free. The acceptance gate for any candidate must compare setup,
ingress, edit_commit, commit_drop, retained-live bytes, type sizes, exact output
ZIP hash, source immutability, no-op behavior, inverse restoration,
stale/foreign-source refusal, cancellation, and finite budget failures. Every
candidate must keep the existing full-parse fallback and source freshness checks.
Do not infer a winner from the lexical byte sum alone.

## ADR and semantic obligations

This direction follows the accepted snapshot model: ADR 0003 requires immutable,
cheap-to-share `Send + Sync` snapshots and atomic commits that leave the source
unchanged (`docs/adr/0003-snapshots-edits-and-patches.md:6-25`). `Arc<str>`
shares immutable numeric lexicals without exposing archive or executor types.
The optimization order and allocation/layout requirements in `docs/GOAL.md:49-57,
97-103,466-496` support removing this measured clone work, while the exact
lexical preservation rule in `GOAL.md:488-489` forbids normalizing numbers to
binary floats or shortened strings.

A successful implementation must preserve:

* exact `Number::as_str()` spelling and all finite/non-finite validation errors;
* complete semantic parse, selected parse, reduced readback, and fallback parse
  behavior;
* every `Stored` metadata field, omission collision refusal, sorted traversal,
  extents, merges, styles, and tolerance records;
* source snapshot immutability and source-version/freshness checks;
* `MultiCommit::snapshot`, `MultiPatch::before/after`, patch application,
  inverse, exact no-op, and output byte identity;
* cancellation, memory/input/output/work/execution budgets and typed errors;
* `Send + Sync` behavior under concurrent immutable reads;
* unchanged source-backed publication and all existing non-XLSX routes.

## Source identities

| input | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell.rs` | `b54c44440100a8cf3729f257ff20c481c896b49a89769d6d5b09e8c14efe370c` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `60d89d8af3847d3df5c96801537e58ec676446fb5d1828b4adbc54d92fc6a462` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |
| `crates/litchi-xlsx/src/cell_values/patch.rs` | `762c222c7fc3580a66c4c1f11f9de44d4a43d4203210597cacdfe34fc17e4648` |
| `crates/litchi-xlsx/src/raw/worksheet/semantic.rs` | `d770235faba92f561e9ed10c1460aedde43fa3fcffbf0423005c6d82843b3268` |
| `crates/litchi-xlsx/src/raw/worksheet/selected.rs` | `f83cefd73e2cd5ab8f741fbd3e9eab9c2ee3793a211f5e2be63cfd7e39af1eb2` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `16983ebda28ace9081be68e6d6a84698f91e95b152e25481c348e6a14d288303` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `8b2079cd3c5437d7e7113e5d4486bed9ed8eb677e87ace82c63b8ce1559f5084` |
| `tools/perf-baseline/src/lib.rs` | `b2ae866c0b10062c57170bd0ef35b5cfd6a46932388718740298134d8272c1be` |
| `docs/GOAL.md` | `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` |
| `docs/adr/0003-snapshots-edits-and-patches.md` | `9200d3546d91f1604d44a9ade7ad0a6bdea50b6442d12a0517b09db5e8e802ef` |
| current source commit | `e29699f49af54e435145cb6051668f6929e7e16f` |
