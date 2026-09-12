# Draft review: XLSX row-owned primary-span arena

Status: draft candidate for fresh measurement. This document makes no speedup
claim and does not authorize production retention. The candidate is based on
commit `67028ab6037ae6eef15af92a0d540285c3c5362c` (`67028ab60`). OLE2/OOXML
remains the active optimization priority; ODF is deferred.

## Source binding

The accepted ADR and source audit map is retained in
`source-binding.json`. Its SHA-256 is
`11927890dbc7d6b02d64cae9c86a1b337e5a44472ca9d85b026cb2d053957b66` and it
binds all 30 indexed accepted ADR files, `docs/GOAL.md`, the scenario
checklist, the workspace manifests, and the bound current source
owners. The relevant unchanged source hashes at the base revision are:

| File | Base SHA-256 | Candidate SHA-256 |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` | `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` | `883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` | `4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064` |

The exported patch is
`row-primary-arena.patch`, SHA-256
`28269e6051f9ee9453ccfd590a0f89269c400529d973a84f0ac908eb6730c5e9`.

## Design

`CellSlot.primary` changes from an owned `Box<[Span]>` to a
`Range<usize>`. `RowSlot` owns one `Box<[Span]>` arena for the row. A pending
row accumulates the spans in a `Vec<Span>` while the scanner is visiting that
row. A pending cell records the arena index at cell start. The scanner appends
the same direct `<f>`, `<v>`, and `<is>` empty-element and close-element spans
at the same events as the existing cell-local implementation. At cell close,
the slot records `primary_start..row.primary.len()`. An empty cell records
`len..len`; an empty row owns an empty span box. At row close, the row arena is
converted to its final boxed slice exactly once.

Both existing `write_cell` call sites pass their owning row's arena. The
writer resolves the cell range with `primary.get(cell.primary.clone())` in
the existing non-empty, payload-replacement branch and then calls the same
`copy_without` implementation with the resulting span slice. Style-only
updates and empty-cell paths do not perform an arena lookup.

## Preservation and error review

- Span values and source byte offsets are unchanged. No source bytes are
  copied into the arena.
- Every span is appended individually in event order. Duplicates and arbitrary
  multi-span ordering remain representable; no sorting, coalescing, or
  deduplication was introduced.
- Empty cells, empty rows, implicit addresses, namespace-qualified tags,
  formula tracking, MCE and unknown-content guards, source provenance, and
  validation remain on their existing scanner and writer paths.
- Address parsing, tag and attribute capture, formula/fusion behavior,
  reparse behavior, and validation code are unchanged. The writer's new range
  lookup is checked and returns a typed invalid-document error if an internal
  layout range is outside its owning row.
- The scanner's prior cell and row state checks remain in the same event
  handlers and retain their existing error text and sequencing for reachable
  malformed states. The new row arena adds no unchecked indexing or unsafe
  code.
- The patch changes only the internal XLSX snapshot model, scanner, and
  worksheet writer. It adds no dependency, public API, archive ownership, or
  format behavior change.

## Tradeoffs and validation scope

The candidate removes one variable-sized span allocation per populated cell,
at the cost of one row-owned arena allocation and a two-`usize` range in each
cell slot. Arena growth can reallocate while scanning a row. The row-owned
arena bounds the pending growth window to the current row, while the boxed
spans remain retained by the full layout until that layout is dropped. The
change is therefore suitable for a fresh allocation and end-to-end
performance experiment, but no benefit is assumed from the representation
alone.

Only static checks were run for this draft: repository `git apply --check` on
the exported patch and `rustfmt --check` with the repository formatter
configuration on the three temporary candidate files. No Rust build, test,
benchmark, profile, allocator capture, or native capture was run. The
temporary candidate source was held under
`/tmp/litchi-goal-0526-candidate` solely to produce and inspect the patch; it
was removed after the handoff; root independently confirmed its absence.
