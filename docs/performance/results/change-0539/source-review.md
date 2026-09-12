# 0539 candidate source review

The delegated review found no blocking lifetime or semantic issue in the 0537
draft. `Attributes<'a>` yields attributes borrowing the current start element,
and the pinned quick-xml normalization API returns `Cow<'a, str>`. The reference
is consumed by coordinate parsing before advancing the event; numeric fields
are consumed immediately by their integer parser. No borrow is retained in
parser state or snapshots. The cell type explicitly becomes owned before
assignment to `PendingCell::cell_type: Option<String>`.

The full checked scan, encounter-time reference decoding, style→cm→vm validation
order, later cell-type decoding, namespace filtering, duplicate checks,
normalization and error mapping remain intact. The expanded baseline test
exercises plain and entity-decoded numeric attributes with exact stored-field
assertions. Existing tests exercise coordinate/type entity decoding, duplicate
and malformed precedence, formulas, unknown types and retained values across
events. Both baseline and candidate pass the same 1,292 XLSX tests.

The pinned quick-xml API keeps owned attribute input and normalization results
owned. That owned-input lifetime behavior is supported by its source contract
and Rust's type checking; the added behavioral test specifically exercises the
borrowed versus entity-decoded value paths through the worksheet parser. It
does not assert a private `Cow` variant or add a test-only runtime API.

The review recommends keeping the frozen native thresholds despite the small
historical instruction share. Only ownership work inside the decoder is
removable; its whole inclusive cost is not an attainable saving estimate.
Planning call-count reduction is measured separately, planning bytes must not
increase, and incremental allocator peak is not RSS. Performance admission and
the remaining applicable quality checks are separate from this static review.
