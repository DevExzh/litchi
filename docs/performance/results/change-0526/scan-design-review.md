# 0526 XLSX scan/layout source audit

scope: read-only audit of the accepted 0525 XLSX scanner and its consumers

performance_claim: none

disposition: identify two separately measurable follow-ups; retain no
production source change in this audit

This review follows the 0525 queue instruction to inspect the scan_with_limit
owner. It does not revive the rejected 0514/0516 emitted event fusion or the
rejected 0522 cell-reference/tag fusion. OLE2 and OOXML remain ahead of ODF,
and iWork is outside this review.

## Evidence and source identity

The audit was performed against commit
67028ab6037ae6eef15af92a0d540285c3c5362c and the source binding in
source-binding.json. The following hashes identify the source and profile
evidence read here.

| artifact | SHA-256 |
| --- | --- |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs | 137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8 |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs | 71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs | 71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5 |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/wire.rs | 86418ecd5c0517d6c5dd000c1e126db51426b3231d0a96d1bf31bc80459c707f |
| crates/litchi-xlsx/src/raw/worksheet/edit/package.rs | b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035 |
| crates/litchi-xlsx/src/raw/worksheet/edit/validation.rs | 87c95be7f6730acd24ba86a349c990e3cb9e7a888b863107c85cf4577cba70ee |
| crates/litchi-xlsx/src/raw/worksheet/edit/codec/validation.rs | 038b4fab325f54e576d4c80f96f7142ec81207f9869ffe60cdca92f7fd0fe491 |
| docs/performance/results/change-0525/candidate/profile-analysis.json | 81cc89e5557fbe538f1344b551284999b6131d3de2da5840533e9a6b658899b5 |
| docs/performance/results/change-0525/candidate/profile-r1-dense-sparse.inclusive.txt | b641d0996f107da9459fba061a3c624c98805d95d2adee2c5c69c7d02c378b8a |
| docs/performance/results/change-0525/next-priority-review.md | bbbe33a47768021bc22606e413853808b3dba73991634e747dc45939fdbd26ab |
| Cargo.lock | 9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a |

The profile is Callgrind instruction attribution from the 0525 candidate binary.
It ranks work; it does not establish a native latency, allocation, RSS,
scaling, or I/O result. No build, test, capture, or allocator run was
performed for this audit.

## Scanner data path and contracts

scan_with_limit in scan.rs:200-317 reads every quick-xml event from the
worksheet byte slice. It checks the event count, uses NsReader with
check_end_names = true, obtains source positions, resolves element namespaces,
and dispatches to the stateful Scanner. Start and empty elements are classified
by parent frame and expanded SpreadsheetML name. The scanner records source
spans and owned tags for the root, defaults, columns, rows, cells, dimensions,
and merge records. It also collects protection, data-validation, formula, and
merge compatibility state.

At scan.rs:952-1002, start_cell calls cell_address and then cell_tag.
cell_address preserves explicit and inferred row/cell coordinates, checks that
an explicit A1 row belongs to the current row, rejects grid overflow, and
updates the implicit-column cursor. cell_tag in wire.rs:98-148 preserves the
source qualified name and decoded attributes while using the Option<Tag> niche
for the common unprefixed cell form. The attribute walk and its error order are
therefore part of the scanner contract.

Known direct SpreadsheetML f, v, and is children are captured as source spans.
The current PendingCell at scan.rs:48-57 owns a Vec<Span>; on cell close at
scan.rs:836-857, that vector becomes a Box<[Span]> in CellSlot
(model.rs:25-38). mce_payload is retained separately. An unknown direct cell
child, including an MCE AlternateContent, marks the cell as incompatible so
ordinary cell validation blocks an edit. An extLst child is allowed to remain
opaque. Primary spans are pushed in XML event order, including repeated or
empty primary elements.

The resulting Layout is consumed by three package routes:

| consumer | scanner state it needs | primary-span use |
| --- | --- | --- |
| package.rs:43-156, rewrite | tags, spans, rows/cells, dimension, defaults, columns, validation, formula, merge and protection state | write_sheet_data reaches write_cell for payload updates |
| package.rs:163-257, rewrite_value_only_with_provenance | the same layout and tags, plus exact row/cell spans | write_sheet_data_with_provenance reaches write_replacement_row and write_cell |
| package.rs:298-425, rewrite_merges | protection, merge records, formula ranges, insertion position, dimension and compatibility | none |

The other layout consumers are edit/validation.rs, which checks protection,
data validations, group formulas, covered merges, dimensions, and MCE payloads,
and codec/validation.rs, which chooses extension prefixes from retained
root/sheet/default/row tags. The semantic raw worksheet parser is a separate
parser and must remain the authority for semantic readback. No scanner event
may be passed into that parser as a replacement for its validation or parse.

The only actual primary-span consumer is copy_without at
write/sheet_data.rs:713-720, reached from write_cell at
write/sheet_data.rs:489-543. The two callers of write_cell are the ordinary row
writer and the value-only replacement-row writer. copy_without expects its
spans to be sorted, non-overlapping source ranges and copies every byte between
them. That behavior preserves whitespace, comments, processing instructions,
extLst, and other opaque bytes around replaced primary payloads.

## Finding 1: potentially removable namespace lookup on non-start events

The loop at scan.rs:214-221 calls reader.resolver().resolve_event(event) for
every event. Scanner::start and Scanner::empty use the returned namespace; the
Event::End branch at scan.rs:260-265 ignores it, and text, CDATA, references,
comments, declarations, processing instructions, and doctype events are also
not namespace-classified by this scanner.

The pinned quick-xml 0.41.0 implementation in Cargo.lock makes the split safe
to investigate. NsReader::read_event_impl calls its internal process_event
before returning: that method pushes namespace declarations on start/empty
elements, schedules the empty-element pop, and schedules the end element pop.
NamespaceResolver::resolve_event itself only computes a ResolveResult; it does
not maintain the scope. For an end event it only performs resolve_prefix on the
literal closing prefix. The reader's check_end_names still performs the
existing literal start/end-name check.

The scanner can therefore be given ResolveResult::Unbound for events other than
Start and Empty, while retaining the existing resolver call for those two
variants. This removes a lookup that has no consumer and also avoids the
temporary unknown-prefix Vec that resolve_prefix can create for an ignored end
name. It does not skip XML parsing, end-name checking, namespace push/pop,
source-position accounting, event limits, or scanner frame finalization.

This is a separately measurable candidate and must not be bundled with the
span-layout change. The dense r1 profile attributes approximately 10.4 million
instructions to the whole NamespaceResolver::resolve_event edge in the scanner
path. That edge includes the required Start and Empty lookups, so the profile
does not quantify the removable non-Start/Empty subset. Its exact ROI cannot be
ranked against the span-layout candidate from this evidence; the subset remains
unquantified until an isolated source-bound measurement.

Before any implementation, a focused differential test must cover:

* default and prefixed SpreadsheetML roots and children;
* nested prefix rebindings followed by a new start element, proving that
  NsReader scope maintenance is unchanged;
* empty elements and their immediate following start element;
* matching and mismatched end names with check_end_names enabled;
* an unknown closing prefix whose result is ignored by the current scanner;
* text, CDATA, comments, processing instructions, declarations, and doctype
  events between recognized elements; and
* identical Layout outcomes and typed errors against the current loop.

The implementation should keep the quick-xml version binding visible in the
candidate evidence. If the dependency changes, the NsReader push/pop and
end-name behavior must be re-audited before reusing this proof.

## Finding 2: per-cell primary-span allocation is removable representation

The current representation performs a heap operation for each cell that has a
recognized primary child. The dense r1 profile reports the following scanner
rows:

| attributed row | instructions |
| --- | ---: |
| scan_with_limit | 137,497,129 |
| Scanner::start_cell | 42,181,691 |
| Scanner::cell_address | 26,816,126 |
| wire::cell_tag | 14,404,797 |
| __rustc::__rust_realloc below the scanner | 5,130,593 |
| RawVec::grow_one below the scanner | 2,522,494 |
| drop_in_place<Layout> | 3,650,816 |

The realloc row is consistent with PendingCell.primary growing for a single
common v, f, or is and then being converted to an exact boxed slice. The
realloc and grow_one rows together attribute 7,653,087 instructions to
allocation-related work, while Layout destruction attributes 3,650,816; none
of those totals is wholly removable because the profile also includes other
collections and retained layout ownership. The attribution does not prove that
every call is solely a primary vector, so a candidate must confirm the change
with an allocator run and a source-bound profile. The row arena remains the
selected pilot representation, without a claimed ranking or measured ROI.

The preferred design is a row-owned primary-span arena:

1. Add a Vec<Span> to PendingRow and a checked Range<usize> (two usize offsets)
   to CellSlot. Keep source offsets as usize; the scan has no 32-bit byte-size
   guarantee that would justify truncation.
2. At cell start, record the current row-arena length in PendingCell. For an
   empty cell, record the same start and end without adding a span.
3. For a recognized primary Empty event, append its exact Span immediately. For
   a recognized primary Start event, retain the existing frame start and append
   its exact Span only when the matching close is finished. At cell close, store
   the range from the recorded start to the current arena length.
4. At row close, move the arena into a Box<[Span]> owned by RowSlot. Rows with
   no primary spans retain an empty slice with no heap allocation. The arena
   grows only for the current row, rather than creating one heap allocation per
   cell or one potentially large whole-sheet buffer.
5. Change both row writer paths to pass the row arena to write_cell. Resolve
   each cell range through a checked slice helper and return the existing typed
   internal invalid-structure error if an impossible range is observed. Feed
   that slice to the unchanged copy_without logic.

The range is sufficient because primary spans are appended while the scanner
visits one row serially, and all primary children of one cell are contiguous in
the arena. It represents zero, one, or arbitrarily many spans without losing
duplicate f, v, or is children. The source offsets and their ordering remain
exact, so copy_without still removes all and only the recognized primary spans.
Bytes between those spans remain copied from the source.

This layout change does not omit scanner work. It keeps every event, namespace
classification, attribute decode, formula observation, MCE flag, source span,
implicit-address check, and resource limit. It also leaves the separate
Tag/mce_payload representation intact. In particular, an unknown direct cell
child that causes a markup-compatibility refusal must not be folded into the
primary arena or made editable by this optimization.

The row arena is preferred over a whole-sheet arena because its peak scratch
capacity is bounded by one row and its ownership follows the existing RowSlot
lifetime. An inline enum or SmallVec<[Span; 1]> would avoid the common one-span
heap allocation, but it would enlarge every CellSlot and can increase retained
layout working-set and copy cost. It remains a possible measured alternative if
the row range changes prove too expensive. A single Span field is rejected
because it cannot represent duplicates or arbitrary primary order.

Selective primary capture based on the edit map is deferred. The plan is known
in the value-only package route, but scan is also used by ordinary
row/column/default rewrites and merge rewrites, and a mode-dependent scanner
would need a complete proof that every possible writer call has retained its
spans. The audit therefore preserves full scan work and removes only the
per-cell storage representation. A selective mode may be considered in a
later source-bound experiment after the row-range contract is established.

## Required correctness coverage for the span candidate

The candidate focused tests must compare the old and new writer outputs and
scanner state for each of these cases:

* a plain empty cell, a cell with one v, and a cell with one f or is;
* duplicate and mixed primary children, including both start/end and empty
  forms, in source order;
* primary children separated by whitespace, comments, processing instructions,
  text, and an opaque extLst child;
* a markup-compatibility or unknown direct child that still causes the same
  typed edit refusal and retains mce_payload;
* inferred row and cell addresses, explicit addresses, prefixed and default
  namespace aliases, and namespace rebindings;
* formula metadata, shared/array/data-table fallback decisions, dimensions, data
  validations, merges, and protection; and
* style-only, clear, payload replacement, remove, new-cell, and no-op routes.

The writer helper test should construct a deliberately malformed private range
and verify a typed refusal rather than a panic. Differential output tests
should assert that replacing a payload removes every recorded primary span while
preserving all other bytes. The existing complete XML validation, semantic
readback, source identity, cancellation, byte/aggregate bounds,
calculation-chain invalidation, and publication tests remain required.

## Rejected or out-of-scope work

The following scanner work was inspected and has no removable proof in this
audit:

* cell_address and cell_tag each serve different contracts. The 0522 candidate
  attempted to combine their common attribute work and failed its
  useful-repeatable native-total gate. It is not requeued under a new name.
* scan_guard collects protection and validation dependencies; removing it would
  change edit blocking. observe_merge_position records schema-safe merge
  insertion and compatibility state; removing it can place a new mergeCells
  record at an invalid location or edit opaque markup.
* Formula scanning and shared_formula_groups establish dependency closure,
  origin/member coverage, duplicate ownership, and bounded shared-formula
  behavior. They cannot be replaced with writer events or a semantic parser
  shortcut.
* tag and cell_tag retain qualified names, namespace declarations, decoded
  values, attribute order, and malformed-attribute error precedence. Skipping
  those checks would weaken lossless preservation and validation.
* Complete output validation and independent semantic readback are outside this
  scanner optimization. ADR 0003 and the 0525 accepted candidate require
  atomic publication and a full validation/readback boundary; this audit gives
  no basis to remove or fuse either pass.

The relevant accepted constraints are ADR 0001's measured-performance and typed
raw-layer boundaries, ADR 0003's immutable source and atomic commit rules,
ADR 0005's allocation/budget measurement contract, and ADR 0006's preservation
and validation rules. The proposed row arena is internal to the XLSX raw codec
and introduces no public API, archive dependency, unsafe code, ambient
execution, or format-ownership change.

## Next action

Measure the two findings as independent candidates. The namespace lookup
candidate should be a small source patch with its own correctness and
Callgrind/native/allocator evidence. The row arena candidate should then use the
same source-backed medium and dense-sparse primary shapes, with duplicate
primary, opaque-interleaving, formula, style-only, and noncompact guards.
Require useful repeatable native total and commit improvements, preserve every
adverse timing/RSS result, and compare allocator calls, bytes, reallocations,
and incremental peak in the canonical lane. Do not infer wall-time benefit from
the Callgrind rows or accept a lower allocation count without an end-to-end
result.

This audit authorizes neither production retention nor a performance claim.
