# 0550 XLSX source-backed commit boundary review

status: bounded read-only review of the current source

production_change: none

This review examines the rewrite/layout, complete validator, reduced readback,
semantic merge, and fallback boundary selected by `change-0549/next-target.md`
against the current source. It is a source review for the 0550 attribution
campaign; it does not claim a hotspot, a removable percentage, or a speedup.
No Rust source, build, capture, fixture, or generated profile was changed by
this review. OLE2 and OOXML remain the active optimization priority. ODF stays
deferred until the OLE2/OOXML goal is complete, and iWork is outside this
scope.

## Identity and prior decisions

The frozen 0550 plan identifies the current source-backed
`MultiSourceEdit::commit` path as the measurement target. Its source revision
is `090b15b64ae52da2f8bf765cbb745ef76122792e`; the checkout has no production
Rust diff. The campaign files are bound by these hashes:

| artifact | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0550/plan.json` | `3eeb52cdad1312b3953006e538490d6523a707e9858354dcb6688b4a615471a7` |
| `docs/performance/results/change-0550/frozen-inputs.json` | `113d7f263fd64d61fc9462d545d33038887009c8b05a05e7ac39afe421376845` |
| `docs/performance/results/change-0550/adr-manifest.json` | `fefccee2e05a78afce65ac8562658a3654fd60895cdcf5701e874e42146ac692` |
| `docs/performance/results/change-0549/next-target.md` | `5b9e75255cfe7010d4ca1bb9426ff389e235a76034d093fb5d14d7c58b930ae4` |
| `docs/performance/results/change-0525/next-priority-review.md` | `bbbe33a47768021bc22606e413853808b3dba73991634e747dc45939fdbd26ab` |
| `docs/performance/changes/0525-xlsx-unchanged-cell-readback.md` | `889e88bcde76886efad653550cf41cfd89466ee309288f709cbf14b2f73a20d6` |
| `docs/performance/changes/0527-xlsx-row-primary-arena-pilot.md` | `c0e89e0b32303b3205d6a3123a023eaf2f61283dc07287671d1e210c06c93aeb` |
| `docs/performance/results/change-0527/source-review.md` | `a0ca6257e870d58397789dc9d0f6bffb6e4eecd47cea325a0ad72fdc059b85cc` |

The 0550 ADR manifest reports that all 30 previously read accepted ADR and
index documents are unchanged from the prior 0549 manifest. This review takes
those retained ownership, immutable-snapshot, validation, resource, physical
package, calculation-chain, and evidence constraints as read; it does not
reopen their architecture decisions.

The relevant current source files are pinned below. The hashes are included so
that any future candidate can be checked against this exact source review,
rather than against a moving working tree.

| source file | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `6db3c616e0e639b578b9aea2493cf717083f4d06f072a201a787002a9daa2178` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |
| `crates/litchi-xlsx/src/cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |

The accepted 0525 change is the current provenance design. It reuses only
unchanged, source-owned cell records after an independent parse of the
changed output. Complete rewritten XML validation, actual output bytes,
independent readback, source identity, fallback behavior, calculation
invalidation, and publication remain part of that accepted contract. The
accepted design also closes structural cases conservatively: an existing row
whose membership changes is retained and parsed as a complete row, while
other unchanged owners may still omit proven source spans. New rows, implicit
later-row followers, shared-formula closure, unsupported metadata,
markup-compatibility content, and any proof failure make the provenance proof
ineligible and force a complete worksheet parse.

The 0527 row-primary-arena pilot is a closed rejection. It moved primary
spans from per-cell boxes into a row-owned arena and changed the scanner and
writers to use ranges. Its dense-sparse repeat-2 total p50 and mean reductions
were 1.8214% and 1.8105%, below the frozen 2% end-to-end gate, so the
production representation was restored. The current source has the original
per-cell `Box<[Span]>` ownership. No future source review or profile should
revive the arena under a different name; lower allocation counts or a better
inner edge are insufficient without the required end-to-end result and all
correctness/resource gates.

## Current execution boundary

`MultiSourceEdit::commit` is in
`crates/litchi-xlsx/src/cell_values/source.rs:1126-1232`. For each selected
sheet it first checks execution state and expanded-action limits. An untouched
sheet is cloned with its source XML still bound. A changed sheet is rewritten,
read back, and appended to the candidate sheet list. The workbook is
calculation-invalidated once after all changed sheets have succeeded, then
the candidate is rebound and returned as a `MultiCommit`/`MultiPatch`. The
separate `publish_multi_commit_to_stream` caller writes that committed patch
to a sink; `MultiSourceEdit::commit` itself does not publish a sink. The
single-sheet `SourceEdit::commit` at approximately lines 818-872 follows the
corresponding one-sheet candidate/patch path.

Both commit methods remove exact no-ops before this boundary. `actions.is_empty()`
returns a patch whose before and after snapshots are the same source snapshot,
with no worksheet rewrite, output validation, reduced readback, calculation
invalidation, or changed-action count, preserving the empty patch and source
sharing. In the multi-sheet path, an empty staged sheet likewise clones the
source and does not enter the rewrite/readback route. A separate
`publish_multi_commit_to_stream` call remains governed by its own stream
publication contract. Any optimization that changes the commit behavior
changes a publicly observable lifecycle and is out of bounds.

For a changed eligible worksheet, the current path is:

```text
MultiSourceEdit::commit / SourceEdit::commit
  -> rewrite_value_only_with_provenance
       -> edit scan(content)
       -> validate actions and structural eligibility
       -> write_sheet_data_with_provenance
  -> Snapshot::from_rewritten_value_source
       -> validation::worksheet_xml(complete actual output)
       -> try_rewritten_value_cells
            -> reduced_readback(complete output, omitted output spans)
            -> raw::worksheet::parse(reduced stream)
            -> provenance/entry closure checks
            -> Store::merge_omitted_cells
       -> complete raw::worksheet::parse(actual output) on any refusal
  -> staged cell/value/formula readback checks
  -> calculation invalidation and MultiPatch construction
  -> separate publish_multi_commit_to_stream sink publication
```

`rewrite_value_only_with_provenance` in
`raw/worksheet/edit/package.rs:163-257` always returns the complete ordinary
rewrite bytes. It scans the source to obtain the layout and source spans,
validates actions against that layout, and only records omission metadata when
the value-only proof is eligible. Shared-formula worksheets, shared-formula
payloads, and actions targeting a row that is not already present return the
ordinary rewrite with an empty omission proof. The writer still walks the
worksheet envelope and writes the complete output.

`Snapshot::from_rewritten_value_source` in
`cell_values/snapshot.rs:715-759` first checks that the omission proof is tied
to the exact source slice by pointer identity. Empty or mismatched proofs take
the generic full-output route. For a matching nonempty proof, the complete
rewritten bytes are validated before a reduced buffer is attempted. A reduced
copy, parse refusal, unsupported record, invalid rectangle, or merge refusal
is converted into the ordinary complete parse; speculative errors do not
become new public errors.

## What each pass proves

The passes have different authorities and cannot be considered duplicate just
because they all consume XML-related bytes.

| pass | current source and work | proof it supplies | what source review can conclude |
| --- | --- | --- | --- |
| edit scan/layout | `edit/codec/snapshot/scan.rs:200-317`, reached by the value-only rewrite | source offsets, retained tags/attributes, row and cell addresses, sheet-data layout, formula/shared-formula state, merges, extension and compatibility state, and action eligibility | Required by the current lossless writer. The semantic `Store` has cells and structural metadata but no source offsets, tag spans, or layout. It cannot be deleted on the basis of the raw parser already having loaded the source. |
| complete output validator | `cell_values/validation.rs::worksheet_xml` | independent grammar, namespace/dialect, element-context, attribute, text, root-closure, and markup/dependency checks over the exact emitted output bytes | It is the current output/publication closure authority and runs before speculative reduced parsing. The edit scanner is more permissive and source-oriented; the raw parser has different processing and semantic/resource behavior. No current proof permits deletion or silent error-order changes. |
| complete rewrite | `edit/package.rs::rewrite_value_only_with_provenance` and `write_sheet_data_with_provenance` | exact ordinary output bytes, preserving untouched source slices and applying changed action bytes | Output retention is required. A writer still visits rows/cells and must preserve unknown bytes and source formatting outside changed spans. Any copy/visit reduction needs byte-equivalence and malformed/unsupported coverage. |
| reduced readback copy | `edit/package.rs:262-288::reduced_readback` | a contiguous stream containing the actual changed output plus retained XML structure after omission spans are removed | It gives the independent parser a changed-output authority. The current parser accepts a contiguous byte slice; no source proof permits deleting the copy or replacing it with staged values. A segmented reader or borrowed-span parser would be a new design and must preserve parser and MCE behavior. |
| reduced raw parser | `raw::worksheet::parse` through `raw/worksheet/codec.rs` | independently materialized changed cells, rows, columns, defaults, merges, formulas, limits, and semantic errors from emitted output | It is necessary to verify the changed output before source-owned records are reused. It may be bypassed only for a proven route whose output semantics, errors, resources, and source ownership are equivalent. |
| provenance merge | `cell.rs::Store::merge_omitted_cells:805-852` | combines parsed changed records with only source records named by checked omitted rectangles and rebuilds sorted indexes/extents | This is a concrete possible boundary, but source inspection alone does not show that its cloning, sorting, bounds, and index rebuild can be removed. A private linear merge or index handoff would need exact Store equivalence and fallback behavior. |
| complete-parser fallback | `from_rewritten_value_source` after `try_rewritten_value_cells` returns `None` | canonical semantic parse of the complete actual output whenever the proof route refuses | It is mandatory for unsupported source/parsed metadata, structural changes, allocation/provenance failure, malformed reduced content, source mismatch, and merge collision. Removing it would turn a safe optimization refusal into an observable correctness or error change. |

The edit scanner and complete validator also have deliberately different
error responsibilities. The scanner resolves namespaces and records source
positions for every XML event while retaining enough opaque layout to copy
unknown bytes. It tracks formula groups, merges, row ordering, extension
descent, and markup-compatibility state. The validator independently checks a
conservative value-only XML grammar and context. The raw parser applies
MCE processing and materializes semantic records, then resolves shared
formulas and calls `Store::from_unsorted`. A source-level claim that these are
one parser pass is therefore unproven.

The source `Store` contains sorted stored entries, row indexes, columns,
defaults, merge index, and declared/stored/content/styled extents. Each stored
entry may carry style, shared-string identity, inline-rich data, formula and
shared-formula state, and cell/value metadata. The reduced route currently
requires both source and parsed entries to pass
`shared_string.is_none() && !inline_rich && cell_metadata.is_none() &&
value_metadata.is_none()`. This is an eligibility guard, not permission to
drop those facets from the full route.

## Conditional boundary hypotheses

These are questions for the fresh attribution, not selected patches. The
source review intentionally does not choose among them using historical
percentages or inclusive profiler shares.

### Source-bound layout handoff or cache

Every value-only rewrite currently calls `scan(content)`, even though the
source snapshot already owns a parsed `Store` and source bytes. A source-bound
layout/proof object could, in principle, hand retained offsets and layout to a
later edit, or amortize repeated edits against one snapshot. The current
`Snapshot` does not retain that layout. Adding it would trade source-load
work, memory, source-version/identity checks, and invalidation complexity for
commit work; a one-commit profile cannot establish that it wins. The fresh
profile must isolate scan/layout construction and writer work. Repeated edits
against the same snapshot are required before a cache can be considered.

The handoff must recompute or validate every fact affected by an edit:
dimension and extension planning, row/cell membership, implicit coordinates,
formula and merge closure, unknown/MCE state, source pointer/version, and all
execution fences. It must never reuse stale source offsets against a different
byte slice.

### Output writer traversal and copying

The provenance writer already copies unchanged source spans and rewrites only
eligible replacement cells while retaining the complete output. If fresh
direct attribution shows that output copying or row/cell traversal dominates,
a bounded writer change could coalesce copies or pass a proof-bearing span
iterator between scan and write. Such a change still needs exact output-byte
comparison, preservation of comments/processing instructions/opaque
`extLst`, explicit changed-cell coordinates, and all eligibility fallbacks.
The output cannot be replaced with a reduced stream or staged values.

### Validator/parser handoff

If fresh evidence establishes substantial repeated validation and parsing,
one possible larger boundary is an output observer or a fused validator/parser
that consumes the exact emitted bytes once and preserves the current validator
grammar, error precedence, raw parser semantics, resource limits, MCE handling,
and independent changed-output authority. A callback that merely trusts the
edit scanner is insufficient: the scanner reads source layout and permits
opaque structures needed for lossless rewriting, whereas the validator checks
the emitted value-only grammar. Deleting the validator or treating a parser
success as validator success requires a written equivalence proof and
malformed/error corpus before any implementation.

### Reduced-buffer elimination

`reduced_readback` validates output offsets, allocates a contiguous buffer,
copies every retained gap, and then hands that buffer to the ordinary raw
parser. If fresh counters show this copy is material, a parser API that reads
discontiguous retained slices or a zero-copy span reader could be considered.
It would need to preserve absolute/relative XML offsets as applicable,
namespace resolution, MCE preprocessing, text limits, parser error order, and
the fact that all bytes came from the complete actual output. It also needs a
fallible allocation/resource contract and a complete-output fallback. No such
parser contract exists in the current source.

### Store merge/index boundary

`merge_omitted_cells` first checks ordered one-row omission rectangles and
absence of parsed entries in them. It counts source entries, reserves a new
vector, moves parsed cells, clones source entries, copies merge ranges, and
calls `Store::from_unsorted`, which sorts and duplicate-checks cells and rows,
rebuilds cell-row indexes, recomputes bounds, and creates the merge index.
This is the strongest source-level candidate if fresh direct attribution
shows merge, cloning, sorting, or index rebuild work is substantial.

A private linear merge could consume two already sorted cell sequences,
preserve duplicate exclusion, and rebuild only indexes/bounds that truly
depend on the combined sequence. It might also move structural rows,
columns/defaults, declared extents, and an unchanged merge index from the
parsed store. That is only a hypothesis. It must prove sortedness, duplicate
behavior, omission non-overlap, all Stored metadata, extents, row indexes,
merge semantics, fallible allocation, and exact fallback/error behavior. It
must not use unsafe code or revive the rejected row-primary arena.

### Minor scanner observation

The preceding 0526 audit noticed a possible unused end-event namespace lookup.
It did not establish a safe optimization or a useful end-to-end result. Fresh
line/disassembly evidence may examine it if it is a measurable direct edge,
but it cannot be selected from the old observation and it does not remove XML
event traversal, end-name checking, source positions, or layout closure.

## Required fresh evidence

The 0550 profiles are diagnostic attribution for the exact current
`MultiSourceEdit::commit` owner. They use simulated Ir and one measured
iteration per profile as frozen by the plan; native samples are descriptive
only. Inclusive descendants overlap, so their percentages cannot be summed
into a removable opportunity and historical 0525/0526/0527 percentages must
not select a current boundary.

The fresh profile review should answer these questions for each shape and edit
size:

1. Which direct caller edge invokes `MultiSourceEdit::commit`, and which
   lifecycle/setup/publication calls are outside the exact owner? Record
   symbol identity, inlining, callgraph, and line or disassembly evidence so
   nested inclusive rows are not counted twice.
2. Within the exact owner, what is direct and aggregate attribution for
   `scan_with_limit`/`Scanner::start_cell`/`cell_address`, layout checks,
   `write_sheet_data_with_provenance`/`write_replacement_row`/`write_cell`,
   `validation::worksheet_xml`, `reduced_readback`, reduced
   `Parser::parse`/`finish_parse`/materialization/shared-formula resolution,
   `merge_omitted_cells`, `Store::from_unsorted`, cloning, sorting, bounds,
   merge-index construction, complete-parser fallback, staged readback, and
   calculation invalidation?
3. Does the output take the reduced success route for each eligible shape, or
   fall back? Report route counts and refusal reasons rather than inferring
   eligibility from the corpus name.
4. What are the complete output bytes, omission span count and bytes, reduced
   bytes, copied bytes, parsed cell/row counts, source-cloned entry count,
   allocation calls/bytes, and incremental peak for each route? These counters
   should be collected from the existing bounded instrumentation or a
   separately reviewed diagnostic; they must not silently alter the measured
   production path.
5. Is the suspected work repeated because of actual call edges, or merely
   nested inclusive attribution? A source-level repetition of XML concepts is
   not enough. If the answer depends on inlining, retain compiler/disassembly
   evidence for the exact binary.
6. Does the one-cell versus one-percent shape change the route, omitted-span
   density, reduced size, merge size, or fallback frequency? Does noncompact
   layout or vendor extension content change eligibility and copied bytes?
7. Is repeated editing of one immutable source snapshot common enough to make
   a layout cache worth its load and memory cost? A single commit/save profile
   cannot answer this; a repeated-commit diagnostic would be needed before
   adding retained layout state.

The correctness and boundary matrix must include, at minimum:

* exact no-op and empty staged action sets, proving no rewrite, validator,
  reduced parse, invalidation, or changed count, plus the empty patch and
  source-sharing result; any separate stream publication is checked through
  its own caller contract;
* one replacement on an existing eligible row and one-percent replacement
  across all four frozen shapes, including explicit changed addresses and
  retained untouched bytes;
* insertion, clear, remove, existing-row membership change, new rows, and
  implicit row/cell followers, proving the complete parser route where needed;
* normal formulas, shared formulas, array/data-table cases, formula
  references, and calculation invalidation;
* shared strings, inline rich text, cell/value metadata, styles, merges,
  columns, defaults, stale or expanded dimensions, and the unsupported-entry
  fallback;
* vendor-extension, x14ac, MCE/`AlternateContent`, unknown direct cell
  children, comments, processing instructions, and opaque `extLst` content;
* malformed XML, namespaces/dialect, depth/event/resource limits, invalid
  omission spans, source pointer mismatch, allocation refusal, reduced parse
  refusal, merge collision, and complete-parser error precedence; and
* multi-sheet edits with untouched sheets, mixed changed/unchanged sheets,
  atomic publication, source/version/execution fences, and independent
  semantic reopen/readback.

For any future implementation, retain exact output-byte hashes, source
lineage/identity checks, unchanged-member checks, semantic readback, and
publication receipts beside performance evidence. A native p50/p95/p99 or
allocator delta is useful only after matched before/after runs over the same
matrix and all required gates; an Ir share is not a latency opportunity.

## Invariants for any candidate

The following constraints are source-level blockers against deleting or
fusing work:

* Source execution, cancellation, version, and managed-budget checks remain
  at their existing fences. `MAX_BATCH_EDITS` remains 256 expanded actions.
* Empty plans and exact no-ops remain the current same-snapshot patch path.
  Multi-sheet no-op sheets remain untouched source clones, and calculation
  invalidation occurs once only when a changed workbook is published.
* The complete rewritten output remains owned and is validated before the
  speculative reduced route. Validator-first and complete-parser fallback
  error precedence remains unchanged.
* `OmittedCells` spans refer to the actual complete rewritten output, not the
  source input. They must be sorted, disjoint, nonempty, in bounds, and safe
  under checked subtraction/addition. Omitted rectangles are zero-based,
  one-row ranges with checked row/column endpoint conversion.
* Source reuse requires exact source-slice identity. A proof from another
  source buffer, stale version, or altered source bytes falls back to full
  parsing. Changed cells and implicit/membership-changing cells are never
  omitted; replacement-only rows may omit only proven unchanged cell spans.
* Any unsupported source or parsed Stored entry, MCE/unknown dependency,
  shared-formula closure, new row, reduced-copy/parse refusal, invalid
  rectangle, merge collision, or optional provenance allocation failure keeps
  the complete raw parser fallback. Reduced errors must not leak as new public
  errors.
* The output validator retains transitional/strict dialect and namespace
  checks, text and attribute grammar, XML root/depth/event closure, and
  unknown/dependency rejection. The raw parser retains MCE processing,
  `MAX_XML_EVENTS`/`MAX_XML_DEPTH`, cell/formula/text/encoded-byte limits,
  shared-formula resolution, and semantic materialization.
* Worksheet output and multi-sheet accumulation retain checked effect/output
  bounds, fallible reserves and allocation labels, including the 64 MiB
  `MAX_MULTI_WORKSHEET_BYTES` aggregate bound. A diagnostic counter must not
  relax a bound or convert managed source payload into an unbounded owned
  copy.
* Every Stored facet, row/column/default/merge structure, declared/stored/
  content/styled extent, source payload, and untouched sheet remains
  equivalent. Staged values are checks on the result, not an independent
  output authority.
* Calculation-chain invalidation, source lineage, atomic MultiPatch
  construction and its separate stream publication, untouched-sheet sharing,
  and independent public readback stay unchanged. No public API, dependency,
  or unsafe-code change is justified by this source review.

## Disposition

The current source establishes several candidate boundaries but does not prove
that any pass or copy is removable. The strongest conditional implementation
hypothesis is a specialized, fallible sorted Store merge if fresh direct
attribution proves that merge/cloning/index rebuild is substantial. A
source-bound layout handoff, a segmented reduced reader, and an output
validator/parser observer are larger alternatives that require their own
equivalence proofs and matched evidence. The edit scan and complete validator
cannot be deleted merely because they both traverse XML, and the reduced
readback cannot be replaced by staged values.

No production candidate is selected in this review. Root-owned fresh
MultiSourceEdit profiles should resolve the questions above, with route and
byte/allocation counters and exact call-edge evidence. If no direct boundary
survives those checks, close 0550 without a speculative patch and return to
the measured OOXML/OLE2 backlog. The rejected 0527 row arena remains closed;
ODF work remains deferred behind the OLE2/OOXML optimization goal.
