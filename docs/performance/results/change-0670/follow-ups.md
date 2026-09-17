# Follow-ups and frozen boundaries for change 0670

## Paragraph-index memo: design only

This change does not add the package-level paragraph index memo. The frozen
design for a later priced implementation is:

* The scope is one immutable package/document owner and one visible main-XML
  generation. The memo is `OnceLock<Option<Arc<ParagraphIndex>>>`; clones of a
  snapshot or document view share the `Arc`, and the index stores checked
  ranges and metadata rather than another XML copy.
* The first paragraph query lazily builds the index. An unmanaged owner pays no
  hierarchical admission charge. A managed owner reserves a typed
  `DocumentIndexAdmission` before parsing: index-range memory and metadata,
  index objects, parser work over the visible XML, and the existing depth
  ceiling. The reservation covers the retained `Arc` for the entire cache
  generation and is released when its last owner drops it.
* A failed reservation, cancellation, or limit check returns before publishing
  a partial index or partial query result. An edit creates a new visible XML
  generation and therefore a new memo admission; it never mutates an old
  snapshot's memo.
* The implementation must price the actual range/vector layout and the
  namespace work before it grows state. It must then measure first-query and
  repeated-view costs against the no-memo path. Until that budget and evidence
  exist, the memo remains design text and no performance claim is made.

This records the design requested by change 0651 row 14 without introducing
speculative retained state into the current parser.

## Guarded tail reader

The remaining `read_event_into` calls in
`source_backed/tail_append.rs` and `source_backed/paragraph_copy.rs` stay in
place. They operate on guarded `BufRead`/source-backed paths whose bounded
token window and source-freshness checks are part of the refusal order. The
slice-backed readers changed in this wave use `read_event()` safely; replacing
the guarded readers requires a separate proof that the same token ceiling,
cancellation checks, and no-output-on-refusal order remain intact.

## Body-final section properties

The editor continues to refuse a direct-body `sectPr` followed by another
body child. The local Strict `wml.xsd` defines `CT_Body` as block-level
children followed by an optional final `sectPr`; changing insertion semantics
would require a separate model decision. Existing reader admission and writer
refusal behavior remain as documented by change 0650's witnesses.

## Publication audit dependency

The transaction range and BOM carry fix is in this branch. The independent
xml-minifier slice-audit offset fix is root's change 0677, commit
`029b17b22`; root should retain that prerequisite before treating marked
managed edit publication as an end-to-end result.
