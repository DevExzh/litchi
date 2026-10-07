# 0959 Number storage and parser owner review

## Scope and result

This is a read-only source review of the source-backed XLSX value editor at
repository commit `3fff4906c4e54382728a7fd64a50445e41d826dc`. It incorporates the
exact-stage allocation attribution in 0959 and the earlier 0958 numeric-owner
review. No source edit, build, benchmark, or new workload was run.

The immediate conclusion is to keep `Number` representation work separate from
the current largest parser owner. The 0959 exact selected-stage ranking reports
`Parser::finish_cell` at 540 calls and 102,826,752 requested bytes, while
`Number::new` is 177,504 calls and 993,756 requested bytes and `Number::clone`
is 175,872 calls and 984,570 requested bytes. Those totals are raw requested
allocator bytes in a nearest-owner stack classification; they are not retained
bytes, phase totals, or proof that the parser work is removable. They do show
that changing `Number` alone cannot address the largest currently attributed
owner.

The smallest defensible Number candidate remains direct immutable shared
lexical backing (`Arc<str>`) behind the existing public behavior, with a private
borrowed parser constructor. An inline-plus-fallback representation is a second
candidate only after a type-size and retained-memory measurement. Neither is
selected by this note. The parser owner review identifies a separate bounded
reserve-hint experiment, with strict preservation gates, rather than removing
the complete source parse.

## Evidence identity and limits

The primary attribution inputs are:

* `/home/zhuhe/litchi-performance-0959/analysis/owner-ranking-summary-v3.json`,
  176,734 bytes, SHA-256
  `f50bb1f2ce39f50d5488d9c8a6a95855fcc98dfc526bbd447e58c871b0f359b4`;
* `/home/zhuhe/litchi-performance-0959/analysis/owner-ranking-v3.json`,
  623,430 bytes, SHA-256
  `d7c61c1cb7a21e4b26ba12ca25a131d4249cf2da260d9ace2fd9e9ac80b1e3fa`;
* `/home/zhuhe/litchi-performance-0959/analysis/semantic-owner-ranking-receipt-v3.json`,
  whose selection is the two fully-qualified
  `xlsx_managed_compression::benchmark::{stage_timed_commit,timed_iteration}`
  stages across the six fixed runs.

The ranking is nearest meaningful `litchi_xlsx` owner attribution. Its labels
for allocator growth, copying, merge backing storage, numeric construction and
numeric cloning overlap. In particular, the 102.8 MB `finish_cell` total must
not be added to the `finish_parse` total or interpreted as a 102.8 MB retained
semantic store.

The source files used for the trace are bound by these SHA-256 values:

| file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell.rs` | `b54c44440100a8cf3729f257ff20c481c896b49a89769d6d5b09e8c14efe370c` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `94b5c115d0e44b51e66bc934a874eb51283567759ccf649a69422f3e096ea2bb` |
| `crates/litchi-xlsx/src/raw/worksheet/model.rs` | `3ad94f003d4ccaba668bab7c2fb531e6887a9594b3d6311edaae07a7f03d37d3` |
| `crates/litchi-xlsx/src/raw/worksheet/semantic.rs` | `d770235faba92f561e9ed10c1460aedde43fa3fcffbf0423005c6d82843b3268` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `183ab8ebb334a3d15e5d0312350915a756467fb578826833da17ac0565d6d53d` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `60d89d8af3847d3df5c96801537e58ec676446fb5d1828b4adbc54d92fc6a462` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/facts.rs` | `eae8c8be1d5af80fa097ecbd60a79d79c94aace5d79e8e00c0f3fab2b327a148` |

## Number's current contract and ownership paths

`Number` is a public opaque `Box<str>` wrapper (`cell.rs:388-447`).
`Number::new(impl Into<Box<str>>)` validates the lexical form and retains its
exact spelling. `as_str`, `as_f64`, `Display`, `Debug`, equality, hashing,
`Clone`, finite-number validation, signed zero, surrounding whitespace and
exponent spelling all depend on that lexical retention. `Value`, `Cell` and
crate-private `Stored` derive `Clone`, so cloning a numeric `Stored` clones the
boxed lexical.

The complete semantic worksheet parser calls `Number::new(value)` for an
ordinary numeric or untyped `<v>` (`raw/worksheet/semantic.rs:168-185`). The
selected scanner validates unselected numerics without owning them, but
materializes selected numerics through an owned lexical
(`raw/worksheet/selected.rs:672-703`); its fallback parse has the same ownership
boundary at `:2117-2125`. The reduced value-only readback parses changed
records and `Store::merge_omitted_cells` clones unchanged source `Stored`
records (`cell.rs:893-959`). That last path is the direct numeric clone site.

The 0958 source-derived medium one-percent route has 9,123 omitted numeric
owners and 58,723 source lexical bytes; the one-edit route has 2,303 omitted
numeric owners and 10,933 lexical bytes. Those are exact corpus ownership
counts, not a phase timing share. They explain why the numeric clone path is a
real, bounded candidate even though the 0959 nearest-owner byte total is small
relative to parser vector growth.

## Number representation options

### Preferred first experiment: private direct `Arc<str>` parser ownership

Change only the private backing from `Box<str>` to `Arc<str>`, keep both as a
fat-pointer field, and add a private constructor that validates a borrowed
`&str` then copies directly into an `Arc<str>`. Use it in
`raw/worksheet/semantic.rs` and the selected scanner. This lets
`Stored::clone` share a numeric lexical allocation while preserving exact
spelling, `as_str`, equality, hashing, formatting, validation and `Send + Sync`
immutable snapshots. The public `Number::new(impl Into<Box<str>>)` boundary
should remain until a deliberate API decision; converting a caller's `Box<str>`
to an `Arc<str>` inside that method can create a temporary Box allocation and
then an Arc allocation, so it is not a free public path.

The Arc tradeoff is measurable: every distinct parsed number gains an Arc
control block, while the generated corpus has little lexical repetition. The
candidate must measure `size_of::<Number>()`, `Cell`, and `Stored`, retained
live bytes at source snapshot lifetime, source/edit allocations, and exact
publication output. A lower clone count alone does not establish a win.

### Deferred: inline short lexeme plus fallback

An inline buffer can copy short integer lexemes inside the Number value and use a
heap fallback for arbitrary valid forms. It must continue to accept the full
bounded lexical domain (up to `MAX_CELL_CHARACTERS` at write validation), retain
whitespace, exponent and signed-zero spelling, and keep `as_str` stable for both
variants. The maximum enum variant and alignment can increase `Number`, `Cell`
and `Stored`; it is not valid to assume that every `Stored` grows or that none
does. Measure all layouts before choosing a capacity.

The historical survey ranked short-number inlining low and cited the rejected
0539 transient-ownership candidate as a likely low-impact result. 0539 reduced
short-lived allocation calls by about 13.6% but changed allocated bytes by less
than 0.5% and failed the frozen workflow gates; it was restored. That rejection
covers the previous broad transient-attribute candidate, not the current
`merge_omitted_cells` numeric clone subset. A fresh candidate would still need
paired native evidence and complete correctness coverage.

### Deferred: Store overlay

An `Arc<Store>` plus a bounded sorted overlay could avoid cloning all omitted
`Stored` entries and avoid append-then-sort/reindex work. The edit action bound
is finite, but this is a broader read-path and iterator design. Historical 0556
prepared a checked linear merge candidate; it did not implement an Arc overlay.
The 0587 survey later proposed the overlay and modelled roughly 16,000
one-edit allocations, while 0557 stopped on baseline noise. These records do
not select an overlay or a Number representation from the 0959 ranking.

## Parser owner trace

### `Parser::finish_cell` is a raw-vector growth owner

At `raw/worksheet/codec.rs:1523-1602`, `finish_cell` first takes the pending
cell, checks value/formula/inline exclusivity and limits, repairs or refuses
formula defects according to `Repairs`, decodes inline text, builds an optional
`RawFormula`, then executes:

```text
self.cells.try_reserve(1)
self.cells.push(RawCell { ... })
```

The attribution line `:1587` is the `try_reserve(1)` call. It is not
`Number::new`, not a semantic cell clone, and not by itself evidence that one
allocation occurs per cell. A standard vector growth can request a new backing
allocation and copy the existing `RawCell` records; the raw total therefore
includes repeated capacity growth/copy work. `RawCell` retains borrowed or
owned lexical fields, formula state, metadata and address until the parser has
finished resolving the worksheet.

The shared source traversal used by source-backed admission constructs a
`Parser`, delivers each validated reader event through `Parser::transition`, and
calls `finish_cell` on close (`raw/worksheet/codec.rs:1624-1715`). Unlike the
ordinary `Parser::parse` route, the observer traversal does not enter the
admitted `lane_body`; it keeps ordinary `Parser::cells` while it concurrently
feeds the XML validator and `FactsBuilder`. This explains why the source-backed
trace reaches the repeated `finish_cell` reserve site.

The 0959 top traces bind this owner to
`Snapshot::from_source_selected` -> `MultiSnapshot::load_source_backed`; the
`finish_parse` trace also appears below `SourceBackedEditor::edit_sheets` in
the benchmark's `stage_timed_commit`. Thus the ranking includes the intentional
selected-worksheet snapshot parse inside the edit setup, not a Number-only
operation.

### `Parser::finish_parse` is final semantic materialization

At `raw/worksheet/codec.rs:628-684`, `finish_parse`:

1. merges repeated rows when the established repair mode allows it;
2. resolves shared formulas before semantic materialization;
3. determines whether a shared-string table is required;
4. allocates the final semantic vector with an exact
   `lane_cells.len() + cells.len()` reserve at `:654`;
5. materializes every retained raw/lane cell, including numeric validation,
   shared-string lookup, formula cache handling, and tolerance records; and
6. passes the vector to `Store::from_unsorted`, which sorts addresses, rejects
   duplicate cells and rows, rebuilds row indexes and extents, and retains the
   tolerance report.

The 32,669,568 requested bytes attributed to `finish_parse` are therefore
primarily final `Stored` backing allocations and related exact vector
materialization. A complete parser cannot simply discard the raw records before
shared-formula resolution, duplicate-row policy, shared-string checks and
error/tolerance decisions have completed. Replacing this with direct
materialization during event delivery would need a separate proof for formulas,
duplicate/out-of-order rows, tolerance ordering and all existing typed errors.

### Why complete source parsing is intentional in the edit route

`SourceBackedEditor::edit_sheets` loads a `MultiSnapshot`; each selected
`Snapshot::from_source_selected` validates worksheet relationships and XML,
parses a complete `Store`, validates style references and scalar cells, and
retains source-bound facts when the shared traversal is admitted
(`cell_values/source.rs:470-475`; `cell_values/snapshot.rs:467-510`). The
semantic store is needed to resolve selector ownership, refuse unsupported cell
owners, preserve source freshness and metadata, and perform staged readback.
`FactsBuilder` additionally retains bounded byte spans used by the value-only
rewriter; it is not a semantic replacement for the Store.

After an edit rewrite, `Snapshot::from_rewritten_value_source` validates the
candidate envelope. When omission spans are present it attempts a reduced
readback and `Store::merge_omitted_cells`; a provenance mismatch, unsupported
cell, collision or reduced-parser refusal falls back to the complete parser
(`cell_values/snapshot.rs:772-859`). That already avoids parsing unchanged cell
payloads on the narrow eligible candidate-output path. The initial selected
source snapshot and every fallback still require complete parse authority.

The benchmark's `stage_timed_commit` calls `editor.edit_sheets` inside the
edit/commit interval, so a source parse in that trace is part of the selected
workflow by contract. The attribution does not justify removing source
freshness, XML validation, complete semantic ownership, or the fallback.

## Smallest bounded parser opportunity

The most concrete parser experiment is a private **best-effort exact reserve
hint** for the eligible shared traversal, rather than a new parsing route.
Ordinary `Parser::parse` already recognizes a strict benign `<sheetData>` body
and reserves `lane_cells` exactly from `lane::Summary.cells`
(`raw/worksheet/codec.rs:451-480`). The observer traversal has the same source
bytes and currently bypasses that lane so that validator and facts events stay
in their existing order.

A narrowly scoped experiment could, when the shared traversal reaches the
`<sheetData>` entry, run the existing pure lane recognizer only as a bounded
hint. If it recognizes the complete body, it can request an exact reserve for
that body before ordinary event delivery. If recognition declines, or if the
hint reservation fails, the current `finish_cell` `try_reserve(1)` path remains
authoritative. The hint must be bounded by the already admitted source/event
limits (`MAX_SHARED_SOURCE_BYTES` and `MAX_SHARED_PROVISIONAL_EVENTS`) and use
checked count arithmetic.

The reserve failure must be ignored as an optimization failure rather than
returned early. Otherwise an allocation error would move ahead of a later
malformed-cell/refusal diagnostic and change historical error precedence. A
successful hint changes allocation timing but not parser events, semantic
checks, or output; the current per-cell reserve remains the fallback for any
body outside the strict lane. The experiment must measure the extra recognizer
scan against avoided raw-vector growth, because a second body scan could erase
the allocation gain.

This is a candidate design, not an implementation decision. A first-pass cell
count, dimension-derived capacity, or unconditional batch reserve is weaker:
`dimension` is producer metadata rather than an explicit-cell count, a first
pass doubles XML work and can alter refusal/limit ordering, and a batch reserve
can move an allocation failure ahead of a later malformed cell. Directly
materializing `Stored` records from `finish_cell` is also not a bounded local
change because shared-formula resolution, duplicate-row last-wins handling,
shared-string lookup and tolerance ordering occur at `finish_parse`.

Required candidate evidence:

* exact semantic-store equality, tolerance report and output ZIP hashes for
  ordinary, source-admitted and fallback parses;
* malformed duplicate attributes, namespace/entity, duplicate/out-of-order row,
  shared-formula, shared-string, overlong value and invalid-number cases with
  the same typed error and precedence;
* source freshness, stale/foreign-source refusal, cancellation, finite input,
  memory, output and work limits;
* allocation counts/bytes for `finish_cell`, `finish_parse`, facts and final
  Store, plus recognizer CPU cost; and
* no change to lane-decline behavior, source snapshot immutability,
  replacement-span authority or non-XLSX routes.

## Decision gate for Number work

Do not start with inline storage or public API changes from the 0959 ranking.
First profile the parser reserve owner and confirm how much of the 102.8 MB is
vector growth/copy versus other overlapping stack categories. If Number work is
still pursued, compare the current Box baseline against direct-parser Arc and
inline-plus-fallback candidates under fresh package/source ownership. Require
before/after type sizes, source snapshot retained bytes, edit/commit allocation
vectors, exact `Number` lexical behavior, source immutability, no-op and
inverse restoration, stale/foreign refusal, cancellation and finite-budget
failures. Keep the complete parser fallback and all existing validation/error
mapping in every candidate.

## Source and history references

* `crates/litchi-xlsx/src/cell.rs:388-447, 646-676, 749-819, 893-959`;
* `crates/litchi-xlsx/src/raw/worksheet/codec.rs:343-394, 451-480,
  628-684, 1523-1603, 1624-1715`;
* `crates/litchi-xlsx/src/raw/worksheet/model.rs:87-183, 211-242`;
* `crates/litchi-xlsx/src/raw/worksheet/semantic.rs:18-134, 159-225`;
* `crates/litchi-xlsx/src/cell_values/validation.rs:266-325`;
* `crates/litchi-xlsx/src/cell_values/snapshot.rs:467-510, 772-859`;
* `crates/litchi-xlsx/src/cell_values/source.rs:470-475, 827-883,
  1160-1235`;
* `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs:246-284`
  and `facts.rs:76-101, 153-181, 410-540`;
* `docs/performance/results/change-0958/review/number-ownership-trace-v2.md`;
* `docs/performance/results/change-0587/survey/xlsx.md:6, 22, 30`;
* `docs/performance/changes/0539-xlsx-transient-ownership-rejected.md:18-35`;
* `docs/performance/changes/0401-xlsx-selected-numeric-elision.md:21-40`;
* `docs/performance/changes/0556-xlsx-sorted-merge-preparation.md` and
  `docs/performance/changes/0557-xlsx-allocation-noise.md`.
