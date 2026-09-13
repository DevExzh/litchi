# 0556 XLSX provenance merge source proof

`status: read-only source proof; candidate preparation`

`performance_claim: none`

`scope: source-backed XLSX value-only commit; OLE2/OOXML priority, ODF
deferred, iWork excluded`

This review audits the proposed private linear merge at
`litchi_xlsx::cell::Store::merge_omitted_cells`. It changes no production Rust,
does not build or capture a candidate, and establishes no latency or
allocation result. The current tree is `3210f07ac7a9daa2686e781645f6539a156a26ce`.

## Constraints and source identity

The relevant `docs/GOAL.md` requirements are to remove measured unnecessary
work, preserve exact output and semantic state, keep bounded fallible
allocations and fallback behavior, and capture before/after evidence before
adoption (lines 50-59, 681-701, 707-739). The 30-entry accepted-ADR manifest
from change 0555 is unchanged: its SHA-256 is
`63b2ce2853c7c548182cb4545cca7c490de914fa4a6791c939518e68a0871b59`, and a
current rehash found zero changed or missing entries.

The source files inspected for this proof have these current SHA-256 values:

| source | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `c684c62aa523cc202027c733c92ad7cba3c91e456b4705c3f7dd9a7876cb53c2` |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `19b4cb00420f895416debe3879f993ad292b1f79973b285b9cc440d6fae11522` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `98a5cf4db40e316cdd58a6904c80bdd11c06f86bd360f0d293651ca521648119` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs` | `b41af5d8c91a5c1e82f9030b8798a354ca4943072373846b8874c73996c03035` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` |
| `crates/litchi-xlsx/src/merge.rs` | `88d005b2a2540bf33fa69bd550c8cb3b7b7d41ae05c2597f55104992f2e0da1e` |
| `crates/litchi-sheet/src/lib.rs` | `d417badda67286fb641148d2189b03e394e2ece1247f490a8bd10ec8ada4d8db` |

## Call path and authority boundary

The changed-sheet path is:

```text
MultiSourceEdit::commit (source.rs:1125-1233)
  -> rewrite_value_only_with_provenance (edit/package.rs:163-257)
  -> from_rewritten_value_source (snapshot.rs:715-759)
       -> complete-output validation
       -> reduced_readback
       -> raw::worksheet::parse(reduced output)
       -> source/parsed eligibility checks
       -> Store::merge_omitted_cells
       -> complete raw parse on any speculative refusal
```

The complete rewritten bytes are validated before the speculative route.
`reduced_readback` only removes checked byte spans from that already validated
output. A failure to construct or parse the reduced view, an unsupported
stored record, an invalid omission rectangle, or a merge refusal returns
`None` from `try_rewritten_value_cells`; the caller then parses the complete
output. `Store::merge_omitted_cells(...).ok().flatten()` deliberately converts
its internal invalid/allocation errors to that same fallback. A candidate must
leave the complete validator, output writer, reduced parser, source pointer
check, execution fences, publication readback, and complete-parser fallback in
place.

Exact no-ops leave `actions` empty in `SourceEdit::commit` or its multi-sheet
counterpart and return the original snapshot/patch. The merge candidate is
reachable only after a real value-only rewrite.

## Proven invariants

`Store` fields are private to `cell.rs`. There are no production struct literals
or mutating methods outside that module. Every production parse constructs a
store through `Store::from_unsorted` in `raw/worksheet/codec.rs:301-323`; the
only other production call is the merge itself. The test helper also uses that
constructor. Consequently both the source store and the reduced parsed store
enter `merge_omitted_cells` through the same constructor invariant.

`from_unsorted` sorts `cells` by `entry.address` and rejects equal adjacent
addresses after sorting (cell.rs:736-753). It sorts `rows` by row index and
rejects duplicates (cell.rs:755-761). It then builds `cell_rows` by transitions
in the sorted cell slice, computes stored/content/styled bounds, and builds a
validated merge index (cell.rs:763-799). `litchi_sheet::Cell` derives `Ord` over
its `row` then `column` fields (sheet `Cell`, lib.rs:332-337), so the cell
ordering is strict row-major order. `Store` cannot be mutated behind the
snapshot after construction; the merge consumes `parsed` and only borrows
`source`.

The derived `Store::default` is another constructor, but it contains empty
cell and row slices (and a default empty merge index), so its sequences are
trivially sorted. There are no production mutations or literals that populate
those private fields outside the constructors above. A test-only malformed
literal can bypass the invariant and is useful for exercising refusal guards;
it is not a production input path.

Therefore a standard two-way merge has a valid precondition for both complete
and reduced stores. A source selection produced by `omitted_entries` is a
subsequence of the source's strict row-major slice. It is sorted provided the
existing constructor invariant is retained. The preflight
`omitted_entries(&parsed.cells, omitted).next()` check rejects any parsed cell
inside an omission rectangle, so every selected source address is disjoint
from every parsed address. An equal address must still be treated as a refusal
if a future helper detects one; it must never silently choose one record.

The parsed row vector is already sorted and unique, and the merge does not add
or remove row records. Moving that vector is therefore safe. `cell_rows` cannot
be moved: omitted source records can change the first cell index for a row and
must be reflected in a rebuilt index. The same applies to
`stored/content/styled` extents. The declared extent must come from
`parsed.extents.declared`, because the source dimension can be stale and the
rewriter may emit a new one.

The parsed columns, defaults, and merge index are authoritative structural
state. The value-only writer has no row, column, default, or merge actions, and
`reduced_readback` removes only cell-body spans, so those parsed structures can
be moved intact if the candidate remains private to this route. In particular,
the source merge index must not be substituted for the parsed one. The current
implementation, and the candidate under review, copy `parsed.merges.as_slice()`
and rebuild an index; a direct move of the already validated parsed
`merge::Index` is a separate follow-up proof question and is not part of this
candidate. Cell records must preserve every
`Stored` field: address, `cell`, style, shared-string identity, inline-rich
flag, formula range, shared-formula storage, cell metadata, and value metadata.
The caller's eligibility predicate rejects shared strings, rich inline text,
and metadata for reuse, but it does not authorize dropping any remaining
formula or style provenance.

## Omission iterator and rectangle proof

`write_sheet_data_with_provenance` records spans in complete-output order. Its
replacement writer walks each scanned row's cells in increasing address order,
uses a sorted `BTreeMap` for actions, and records only contiguous unchanged
runs. Untouched rows record their whole cell body. `record_omitted` requires a
same-row, non-inverted address run and stores offsets into the complete output
(`sheet_data.rs:189-296`).

`reduced_readback` checks that spans are nonempty, in bounds, sorted, and
disjoint before allocating and copying the reduced XML
(`edit/package.rs:260-288`). It does not establish the address proof; the
snapshot caller reconstructs checked `Rect`s with checked row/column addition
(`snapshot.rs:740-750`). `omitted_ranges_are_ordered` then requires every
rectangle to be exactly one row and requires same-row intervals to be
nondecreasing and nonoverlapping (`cell.rs:945-961`). Gaps and later rows are
valid. The row is intentionally part of the previous-end comparison, so a
range such as `D1:E1` followed by `A2` is valid even though the next column is
lower.

`omitted_entries` consumes the source slice once per call. For each sorted
source address it advances past completed ranges, skips rows/columns before a
range, and yields only `Rect::contains` matches (`cell.rs:964-982`). Its two
current callers first count the selected source entries and then clone them.
The linear candidate may combine those passes only after proving that the
selected iterator remains monotonic and that its count exactly matches the
reserved output capacity. It must not use an address interval shortcut that
assumes a rectangle spans multiple rows; multi-row rectangles are an explicit
refusal.

## Error, resource, and fallback obligations

The current merge refusal order is observable through resource behavior even
though speculative errors are swallowed: empty or unordered omissions, zero
selected source entries, and parsed collisions return `Ok(None)` before the
combined allocation. The checked cell-count addition precedes
`try_reserve_exact(total)`, whose resource label is `"merged worksheet cells"`.
The merge-range copy has its own fallible reservation and label. Errors from
either reservation or from `from_unsorted` become `None` at the caller and
fall back to the complete parse. A replacement must preserve checked overflow,
fallible reservation, bounded temporary lifetime, and the absence of panic or
unsafe code. It should keep the current preflight order unless a written proof
shows that resource/error behavior is unchanged.

Do not reuse parsed extents wholesale, trust source declared bounds, suppress
duplicate detection, normalize an equal address, or propagate a speculative
error directly to the public commit. Do not turn a collision or an uncertain
precondition into a partial store. Any temporary merged buffer must be dropped
before the complete fallback parse, as the current `match` does. The complete
output validator and parser remain the public error and preservation
authorities; the merge is only an optimization of already checked state.

## Risk register and required proof checks

The main assumptions that need explicit candidate review are:

1. The candidate must either retain the private-constructor invariant or check
   strict ordering and duplicates before using unchecked index arithmetic. The
   current source proves the invariant, but a future alternate constructor
   would invalidate an undocumented assumption.
2. The parsed sequence must be moved without cloning; only selected source
   entries should be cloned. A source iterator/count mismatch must refuse
   safely rather than panic or leave a hole.
3. Equal or out-of-order addresses must refuse to the complete parser. The
   parsed-in-omitted precheck remains necessary even if the merge loop also
   checks equality.
4. Rows may be moved, but `cell_rows` and all three derived cell extents must
   be rebuilt from the final merged sequence. The declared extent must remain
   the parsed output's value.
5. Moving the parsed merge index is safe only because this private caller does
   not edit merges. A candidate must not accidentally move the source index or
   rebuild from a different range list.
6. All `Stored` provenance fields and structural rows/columns/defaults must be
   compared against a complete parse. Formula, style, explicit-empty, inferred
   address, and maximum-grid coordinates are needed; scalar values alone are
   insufficient.

## Focused test matrix

Before any performance capture, the candidate should pass these checks against
the existing complete-parser oracle:

| area | cases that must be covered |
| --- | --- |
| linear merge | parsed cell before/between/after omitted cells; same-row gaps; adjacent ranges; multiple rows; next row with a lower column; empty rows and explicit empty cells |
| ordering/refusal | empty omissions; unordered ranges; overlapping ranges; multi-row rectangle; no source entries selected; parsed cell inside an omitted rectangle; equal-address refusal and out-of-order iterator guard |
| address provenance | explicit and inferred `r` attributes; untouched rows; changed implicit owners; first/last grid coordinates; stale source dimension with a newly expanded parsed dimension |
| records and indexes | ordinary values, formulas and formula ranges, style, all `Stored` fields, row properties, columns, defaults, merge ranges, covered-merge lookup, cell-row traversal, and stored/content/styled extents |
| caller fallback | malformed complete output; reduced-readback span failure; unknown or markup-compatible elements/attributes; shared formulas; new rows; shared strings; rich inline text; cell/value metadata; source pointer mismatch; source/parsed collision |
| commit semantics | single- and multi-sheet changed commits, clear/remove/readback, exact provenance-vs-ordinary output bytes, inverse patch restoration, and exact no-op source/store sharing |
| resource/safety | checked count overflow, fallible reservation labels, cancellation/execution fences, no panic on refusal, and full workspace lint/format/feature checks |

The existing snapshot tests already cover several important oracle cases,
including same-row gaps with a lower-column next row, scalar facets, implicit
cells, stale dimensions, source mismatch, new rows, shared formulas, malformed
markup, and no-op sharing. A candidate review must retain those tests and add
direct comparisons for any moved row/index/merge state. Allocation failure
claims require an actual bounded allocator or equivalent evidence; source
inspection alone cannot claim them.

## Conclusion

The current private construction graph proves strict row-major, duplicate-free
cell and row inputs for `merge_omitted_cells`, and proves that the omission
iterator yields a sorted source subsequence under its checked one-row range
contract. A checked linear cell merge is therefore source-feasible. The proof
does not authorize deleting complete validation/parsing, reusing cell indexes
or extents, changing fallback/error order, or claiming a speedup. Proceed only
with an isolated candidate that preserves these boundaries, then use fresh
end-to-end and operation-local allocation evidence to decide adoption.
