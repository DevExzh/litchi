# 0959 Number representation review (v2)

- Review date: 2026-10-07
- Checkout: 3fff4906c4e54382728a7fd64a50445e41d826dc
- Scope: read-only review of litchi-xlsx::Number, its materialized and selected worksheet parser paths, and the inline/Arc design direction.
- Result: no source changes, builds, or workloads from this review. The representation candidate is not ready for adoption until the semantic and layout gates below pass.

## Current contract and measured context

Number is a public re-export of the private-field type in crates/litchi-xlsx/src/cell.rs:388-487. It stores a Box<str> and derives Clone, PartialEq, Eq, and Hash. Comparisons therefore use exact lexical contents, not numeric value. Debug prints the Number("...") tuple form and Display writes the unchanged lexical form. as_str returns the stored string directly.

The public constructor at cell.rs:393-398 must remain Number::new(impl Into<Box<str>>) for source compatibility. It first takes ownership, then calls validate_lexical, and retains the original spelling. The integer Value implementations at cell.rs:322-334 construct the private Number(...) tuple directly from a finite integer string; a new representation must keep an equivalent trusted constructor or change these implementations without weakening their invariant.

The 0959 layout probe has no usable size result. Its first attempt terminated with -D unused-qualifications errors before printing any rows (the preserved terminal has return code 101); the second attempt has no parsed output or terminal receipt. The rough Inline15, Inline23, and BoxArc enums in the probe are not production measurements. Do not choose an inline threshold or claim Arc<str> is layout-neutral until an isolated, clean probe reports Number, Option<Number>, Value, Cell, Stored, Formula, and Content sizes and alignments from the exact candidate source.

The 0959 ranking is useful for scope only. In the fully qualified six-run selection, numeric-construction rows account for 177,504 calls and 993,756 requested bytes; numeric-clone rows account for 180,768 calls and 1,012,146 requested bytes, against 2,081,214 selected calls and 274,320,408 selected requested bytes. These labels overlap with XML, copy, and growth labels, so their sum is not an additive saving estimate. They are roughly 0.36% and 0.37% of selected requested bytes respectively. The ranking is allocation attribution and supplies no speed or adoption claim.

## Parser and ownership paths

1. The complete materialized parser calls parse_value in crates/litchi-xlsx/src/raw/worksheet/semantic.rs:168-185. For an untyped or numeric cell it checks value.trim().is_empty(), then calls Number::new(value) at lines 177-184. A malformed/non-finite value either returns the historical error or, when the tolerance report is present, records MalformedNumber and demotes the cell while retaining its raw decoded lexical. A direct parser constructor must preserve this validation, tolerance branch, error text, and order.

2. The plain selected scanner accumulates the decoded value in a reusable String and finishes at raw/worksheet/selected.rs:672-703. Selected cells call Number::new(value.into_boxed_str()) at lines 684-690; unselected numeric cells call only Number::validate_lexical at lines 698-701 and fall back on failure. This path does not perform materialized-parser repairs. Preserve the PlainScanError::Fallback boundary.

3. The general selected scanner finishes at selected.rs:1770-1945. Unselected numeric scratch uses validate_lexical at lines 1896-1905, and selected/formula-cache values go through parse_value at lines 1921-1924. Its numeric branch at selected.rs:2117-2126 calls Number::new(value.to_owned().into_boxed_str()). The scanner later recycles the input String at lines 1948-1951. A direct move optimization must not invalidate scratch reuse, formula-cache ownership, or selected/unselected behavior.

4. raw/worksheet/codec.rs:1539-1543 rejects values over 32,767 characters before materialization. The selected scanner separately charges encoded bytes and decoded characters at selected.rs:758-781; its limits are 32,767 characters and 32,767 * 14 encoded bytes. Public Number::validate_for_write at cell.rs:439-446 is a later character-count check. Inline storage must use these existing limits and must not confuse byte threshold with character limit.

5. Source-backed value edits clone omitted stored records in Store::merge_omitted_cells, cell.rs:893-959, with entry.clone() at line 950. Snapshot::try_rewritten_value_cells calls that merge at cell_values/snapshot.rs:832-858. This is the concrete deep-Number clone path behind numeric-clone attribution. Ordinary snapshot clone in from_rewritten_value_source, snapshot.rs:795-829, is Arc-backed at the surrounding store/source level and is not another numeric constructor callsite.

The production search found no other Number::new callsite in litchi-xlsx; the remaining direct construction is the trusted integer macro above. Tests and StreamingCellValue::Number(f64) are separate APIs.

## Validation hazards

validate_lexical at cell.rs:400-418 has two observable branches. An optionally signed run of at most 18 ASCII digits is accepted by the fast path without a float parse. Every other value is parsed after Rust's Unicode-aware str::trim(); the original untrimmed value is used in invalid/non-finite error text. as_f64 at lines 426-437 applies the same trim and finite filter. The stored value is never trimmed.

Tests must pin this behavior:

- Equal numeric values with different spellings remain unequal: "1", "+1", "01", "1.0", "1e0", "-0", "0", and "  -0.000  ".
- Leading/trailing Unicode whitespace accepted by current Rust trim behavior must be retained byte-for-byte; include U+00A0 and U+2003. Interior whitespace and non-ASCII digits must retain current rejection.
- Entity/CDATA-decoded worksheet values must retain the existing decoded lexical result. Existing selected tests cover fragmented 1.25E+3 and exact -0.000/exponent spellings.
- Empty trimmed numeric values remain Cell::Empty on selected parsing. NaN, infinities/overflow such as 1e999, malformed punctuation, and invalid text must preserve current error or fallback/tolerance behavior.
- Long valid and invalid forms must be tested around the candidate inline boundary and the existing 32,767-character limit. The public constructor can retain a long finite lexical that the write path later rejects by length; do not impose a new constructor limit accidentally.
- as_str, as_f64, Display, and Debug must be panic-free and preserve exact spelling for inline and heap/shared forms.

An inline representation must not derive equality or hashing over its enum discriminant. If the same lexical can be inline in one instance and heap/Arc in another, manual PartialEq, Eq, and Hash must operate on as_str so representations compare and hash identically. Numeric normalization is forbidden. The current type does not promise numeric equality or ordering.

An inline UTF-8 payload needs a real invariant. as_str must not use unwrap, expect, unreachable!, or unchecked UTF-8 conversion on caller-controlled bytes. Construct inline only after checked UTF-8/length validation, keep heap fallback for non-ASCII or over-threshold lexemes, and make the borrow path safe without silently returning empty or truncated text if an invariant is violated.

## Design recommendation

Arc<str> is the lower-risk first comparison for the measured clone path because its pointer-sized representation is expected to avoid the enum payload expansion of a small-string variant, while immutable clones can share lexical bytes. That expectation needs the clean layout probe. The public new(impl Into<Box<str>>) cannot change. A global Arc representation may convert the public Box into a second allocation/copy, so it needs a private parser constructor that validates before constructing Arc directly; do not assume Box<str> to Arc<str> conversion is allocation-free. Measure public Number::new(&str/String/Box<str>) separately from parser construction and clone behavior.

A hybrid inline/heap representation is admissible only if the parser has a direct borrowed/owned constructor. Routing the current public constructor through Box and then copying every short value into an inline buffer adds work and can erase the proposed saving. Its larger Number payload also propagates into Value, Cell, Stored, Content, and Option forms. Choose the threshold from an exact lexical byte/character histogram and layout/peak-memory evidence, not rough 15/23-byte labels. This numeric allocation slice is small; an inline change that increases hot cell-record size or resident memory is not justified by attribution alone.

Any private direct parser constructor must be called only at the three parser paths above, preserve full-parser error/tolerance sequence and selected-scanner fallback sequence, and leave trusted integer construction valid. The immutable shared form must not expose mutation or use copy-on-write in a way that changes source identity or preservation.

## Focused acceptance gates

- Add internal table tests constructing every representation for the same lexical and assert exact as_str, Display, Debug, Eq, and equal Hash; add distinct-spelling cases to prove no numeric normalization.
- Differentially exercise complete materialization, plain selected scanning, general selected scanning, formula cached values, entity/CDATA fragments, empty values, malformed/non-finite values, tolerance repair, and selected fallback. Compare exact errors, defect locations, selected payloads, and source lexemes.
- Exercise Store::merge_omitted_cells with the source-backed value-edit corpus, including no-op, changed values, inverse restoration, source freshness/version checks, limits, cancellation, and malformed input. Verify exact output bytes and inverse restoration.
- Test public constructor inputs of &str, String, and Box<str>; test inline boundary N-1/N/N+1 in bytes and characters, Unicode outer whitespace, long finite forms, overlong write rejection, and allocation/refusal behavior. Keep as_f64 finite/None behavior unchanged.
- Re-run the clean integrated layout probe and record sizes/alignments for all hot containers before and after. Record allocation counts/requested bytes for public construction, parser construction, and deep clone separately; include peak memory/RSS and source lexical distribution.
- Use existing representative source-backed and generic controls and qualification gates. Numeric attribution is a small overlapping slice, so no adoption should follow from allocator rows alone; require end-to-end exact source/output/inverse and preservation/security checks.

## Bound source identities

- crates/litchi-xlsx/src/cell.rs: b54c44440100a8cf3729f257ff20c481c896b49a89769d6d5b09e8c14efe370c
- crates/litchi-xlsx/src/raw/worksheet/semantic.rs: d770235faba92f561e9ed10c1460aedde43fa3fcffbf0423005c6d82843b3268
- crates/litchi-xlsx/src/raw/worksheet/selected.rs: f83cefd73e2cd5ab8f741fbd3e9eab9c2ee3793a211f5e2be63cfd7e39af1eb2
- crates/litchi-xlsx/src/raw/worksheet/codec.rs: 94b5c115d0e44b51e66bc934a874eb51283567759ccf649a69422f3e096ea2bb
- crates/litchi-xlsx/src/raw/worksheet/tests.rs: 44c5db3893ccde5bbbefa838417b0b8f20020effb71c9f6ef276319baa41f4c2
- crates/litchi-xlsx/src/cell_values/snapshot.rs: 50a4ef8fd9f2e810d2b1315d8a389086134a28901a2b86a98cc8abeb84308726
- crates/litchi-xlsx/src/cell_values/source.rs: 60d89d8af3847d3df5c96801537e58ec676446fb5d1828b4adbc54d92fc6a462
- crates/litchi-xlsx/src/raw/worksheet/model.rs: 3ad94f003d4ccaba668bab7c2fb531e6887a9594b3d6311edaae07a7f03d37d3
- docs/adr/0005-io-memory-and-performance.md: 8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa
- docs/adr/0006-validation-security-and-compatibility.md: ae21189eda0acc9524a7c5a57e87bb3566405ede5c015c87105e01337666f44b

This review used no production candidate source and made no workload claim.



## 0959 v2 evidence update: integrated layout and strict callsite selection

The v1 review is preserved verbatim in this file before this supplement, and the
original v1 file remains unchanged (11974 bytes, SHA-256
`06d542ffcdef5d312543087d1dcfadd28e8fcdda0f8844a4c9391943474d547c`). The
clean layout result is now available at
`/home/zhuhe/litchi-performance-0959/receipts/layout-probe-result-v2.json`
(5691 bytes, SHA-256
`601dd859bec08538cf74d0364cb73e5e69416017f6d0ea65aec9a1cc37a51411`). Its
status is `pass`; the captured stdout is
`review/layout-probe-v2.stdout` (1530 bytes, SHA-256
`265847d2b8ff953697f42e9b0ce62b827e0414d88a88e2865eb8eb281c7ab91e`), and the
terminal receipt is `receipts/layout-probe-terminal-v2.json` (1359 bytes,
SHA-256 `1e5bdb8475a8c9b77b2f07434bd80b89dc56845f72b55490af81f1218eca38cf`).

The exact current integrated layouts are:

| type | size | alignment |
| --- | ---: | ---: |
| `Number` | 16 | 8 |
| `Value` | 24 | 8 |
| `Cell` | 72 | 8 |
| `Stored` | 184 | 8 |
| `Formula` | 72 | 8 |
| `Content` | 72 | 8 |
| `Option<Number>` | 16 | 8 |
| `Option<Value>` | 24 | 8 |
| `Option<Cell>` | 72 | 8 |
| `Option<Stored>` | 184 | 8 |
| `Option<Formula>` | 72 | 8 |
| `Option<Content>` | 72 | 8 |

The probe's alternatives remain explicitly non-integrated rough types. They
measured `Inline15` at 16 bytes/alignment 1, `Inline23` at 24 bytes/alignment
1, and `BoxArc(Box+Arc<str>)` at 24 bytes/alignment 8, with the corresponding
`Option` sizes unchanged. These rows do not implement `Number` semantics and
do not establish the layout of an integrated inline candidate. The evidence
does establish that a representation larger than the current 16-byte `Number`
would propagate into at least `Value`, and potentially the 72-byte `Cell` and
`Content` or 184-byte `Stored` records; an Arc-backed `Number` candidate has
preserved the current integrated layout in its isolated probe.

The corrected strict-stage ranking is
`/home/zhuhe/litchi-performance-0959/analysis/owner-ranking-v3.json` (623430
bytes, SHA-256
`d7c61c1cb7a21e4b26ba12ca25a131d4249cf2da260d9ace2fd9e9ac80b1e3fa`), produced
by `tools/rank_0959_v1.py` (22650 bytes, SHA-256
`3daeaac5269257d60afc1089434c7eaf3d130d85ff6e62cb8bf2194adc0be2be`). The
helper's selection predicate is source-visible and exact at the function-path
level: it checks for both
`litchi_perf_baseline::xlsx_managed_compression::benchmark::stage_timed_commit::`
and
`litchi_perf_baseline::xlsx_managed_compression::benchmark::timed_iteration::`
prefixes in the flattened trace locations, and rejects rows carrying either
unknown-frame marker. The trailing `::` intentionally permits Rust's
monomorphization/hash suffixes. It is not the prior substring-token or
first-match resolver. The predicate does not independently verify a source
file/line for each function; that limitation is covered by the sealed input
manifest, reference-analysis hash, function namespace, and source/build
custody, and is why this remains an offline attribution result rather than a
phase-timing proof.

I independently replayed the old and new predicates over the six sealed
`raw-decoder-v1/*/trace-aggregates.jsonl` files. Every one of the 2604 rows in
the former admitted `MultiSourceEdit::commit` bucket also contained both exact
benchmark callsites and no unknown marker. Thus the former surrogate rows are a
strict subset of the corrected selection, not a disjoint population. Their
aggregate remains 788346 calls and 58104786 requested bytes. The corrected
per-run and aggregate rows are:

| run | exact calls | exact requested bytes | exact trace rows |
| --- | ---: | ---: | ---: |
| `h00-medium-percent-r1` | 330372 | 51500040 | 868 |
| `h01-dense-percent-r1` | 628254 | 73812219 | 870 |
| `h02-medium-one-edit-r1` | 81981 | 11847945 | 842 |
| `h03-medium-one-edit-r2` | 81981 | 11847945 | 842 |
| `h04-dense-percent-r2` | 628254 | 73812219 | 870 |
| `h05-medium-percent-r2` | 330372 | 51500040 | 868 |
| **aggregate** | **2081214** | **274320408** | **5160** |

Every run reports conservation `pass`; the six selected totals exactly match
the prior 0958 `owner_selection.selected_owner` totals. `MultiSourceEdit::commit`
is now only a semantic owner/tag, never a selection surrogate. The separate
`owners_ranked` list still chooses the first meaningful `litchi_xlsx` frame
above allocator/runtime frames and its labels overlap; those owner rankings are
not an additive inclusive call tree and do not establish speed, retained bytes,
or adoption.

## Arc candidate source review

The external Arc candidate is at
`/home/zhuhe/litchi-performance-0959/source-candidate-arc-v1`. A byte census
against `source-baseline` finds exactly three changed production Rust paths:
`crates/litchi-xlsx/src/cell.rs`,
`crates/litchi-xlsx/src/raw/worksheet/semantic.rs`, and
`crates/litchi-xlsx/src/raw/worksheet/selected.rs`; the candidate also carries
an external `Cargo.lock` for its isolated build. The candidate layout probe
passed with the integrated sizes above. Focused semantic tests were reported as
passing by the quality owner; a final release-build receipt and comparative
workload are not part of this review.

The source preserves the public `Number::new(impl Into<Box<str>>)` signature,
validates before storing, retains exact lexical spelling, keeps the existing
`validate_lexical` error/tolerance behavior, and derives content-based
`Eq`/`Hash` over `Arc<str>`. The private `from_lexical(&str)` constructor is
used by the complete semantic parser and the selected scanner, so the parser
paths do not borrow XML data after return. The selected plain path still maps a
failed numeric validation to `PlainScanError::Fallback`; the complete parser
still records `MalformedNumber` only through its existing tolerance branch.
The trusted integer conversions now build an `Arc<str>` directly from itoa's
stack lexical, preserving the old infallible integer proof without adding an
`expect` panic path. The current cell source hash after that correction is
`f345c0ab2101d5dad2b25da4bfbfcb3f1bdafabdeca2ad31f66535887e5c2291`; the old
`layout-arc-v1-source-end.sha256` receipt still records the pre-correction
`b066aad...` hash and must be regenerated before any release-build custody is
accepted.

One candidate-specific measurement obligation remains. In the plain selected
scanner, `finish_cell` owns a reusable `String`; the baseline moves that
allocation into `Box<str>` when a selected value is materialized, whereas the
candidate passes `&value` to `from_lexical`, allocates an Arc copy, and then
allows the selected scratch `String` to drop. The general selected parser
recycles its scratch after parsing, so its ownership tradeoff differs. The
candidate therefore needs allocation and end-to-end checks for both selected
routes before it can be called a no-regression change. The public Box-to-Arc
conversion in `Number::new` also must be measured separately; its API remains
compatible but it is not assumed allocation-free. No source or production
adoption follows from the unchanged layout or focused semantic tests alone.

## v2 disposition

The corrected trace selection and layout evidence pass their bounded offline
checks. The Arc source is semantically coherent for the reviewed parser paths,
subject to refreshing its source receipt and measuring the plain selected
ownership tradeoff plus full quality/publication gates. This v2 review makes no
production edit, build, comparative workload, speed claim, or adoption claim.
