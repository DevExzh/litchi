# 0537 worksheet parser source review

**Disposition:** bounded candidate for fresh validation; diagnostic only. No
production optimization is admitted by this review, and no fresh timing or
allocation result is claimed.

This review is bound to baseline revision
`97b75c24c4a92098458972ca94f183ff6d256bac`, baseline
`raw/worksheet/codec.rs` SHA-256
`69719bc7a0aa0754ab4745f49baba077303ca7909f4e9adbc5243189e85fb9a7`, and
candidate patch SHA-256
`5615b3433d83b960167c442e7fb7c13fe3a295fc2ac78de859c049af12f396aa`.
The source custody record is the [source binding](source-binding.json),
SHA-256 `ae8983fdaf023fbf1145a468a468c15d54f48ea98d70ecaa57f7d5c260d76416`.
The descendant replay is [analysis.json](analysis.json), SHA-256
`affb2bc67f7794cdb898ab54982fabae0732aeca0a1a669d994dc2746d596121`, under
[plan.json](plan.json), SHA-256
`261becdadcaa34373d5619f508fec4daffbdc924d61a595533aafd02fdb153b2`.

## Owner boundary and evidence

`Snapshot::from_source_selected` performs the required execution check, sheet
kind and relationship refusal, package part read, `SourcePayload` retention,
aggregate byte check, complete worksheet XML validation, raw worksheet parse,
post-parse execution check, style-reference validation, and scalar-cell
validation ([current source](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L426)).
The original worksheet bytes remain in the resulting source state. The raw
parser is therefore an owned semantic-parser step inside a larger immutable
source snapshot; it cannot be removed, skipped, or replaced with a partial
cell scan under the current contracts.

`raw::worksheet::parse` still performs extension capture when needed, MCE
processing, UTF-8 validation, and the complete event parser, including rows,
columns, formulas, shared strings, inline strings, unknown cell types, and
resource-bounded store construction ([current source](../../../../crates/litchi-xlsx/src/raw/worksheet/mod.rs#L28)).
The retained 0530 planning review measured this complete raw parse at roughly
60% of the selected planning owner, but that inclusive parent attribution does
not establish that the small attribute helper is removable work or that this
candidate has a corresponding end-to-end gain. See the [0530 planning source
review](../change-0530/planning-source-review.md).

The sealed 0530 Callgrind descendants, re-bound to the current source, are:

| corpus/repeat | planning Ir | `scan_cell_attributes` inclusive Ir (share) | `decode_cell_attribute` inclusive Ir (share) | direct `__rust_alloc` Ir |
| --- | ---: | ---: | ---: | ---: |
| medium/r1 | 125,643,633 | 10,273,383 (8.1766%) | 3,871,719 (3.0815%) | 1,584,480 |
| dense-sparse/r1 | 237,075,784 | 20,032,877 (8.4500%) | 7,601,777 (3.2065%) | 3,031,751 |
| medium/r2 | 125,655,087 | 10,273,571 (8.1760%) | 3,871,907 (3.0814%) | 1,584,668 |
| dense-sparse/r2 | 237,102,562 | 20,034,783 (8.4498%) | 7,603,683 (3.2069%) | 3,033,657 |

`decode_cell_attribute` is nested in `scan_cell_attributes`, which is nested
in planning; inclusive rows must not be added. The `__rust_alloc` column is
direct Callgrind instruction-reference attribution, not an allocation count
or operation-local byte total. These are historical 0530 profiles, not fresh
0537 measurements. The [descendant analysis](analysis.json) records the raw
rows and this nesting explicitly.

## Candidate shape

The four-hunk candidate is confined to the private raw worksheet codec
([current source](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs#L172)):

* Import `Cow` and change `CellAttributeView::reference` to
  `Option<Cow<'a, str>>`.
* Make `decode_cell_attribute(Attribute<'a>, Decoder)` return the crate's
  `Result<Cow<'a, str>>` from quick-xml's normalized value, retaining
  `XmlVersion::Explicit1_0` and the existing malformed-value mapping.
* Let the reference be borrowed through `parse_a1` and let style,
  cell-metadata, and value-metadata use the returned text for their immediate
  `u32` parse. The existing `Cow` dereference gives the same lexical text to
  `parse` and to error formatting.
* Convert the decoded `t` value with `Cow::into_owned()` before assigning it
  to `PendingCell::cell_type`. That field, and the later `RawCell` field, are
  `Option<String>` ([model source](../../../../crates/litchi-xlsx/src/raw/worksheet/model.rs#L87)),
  so `t` cannot borrow the event.

The checked one-pass scan, unqualified/prefixed filtering, raw deferred
attributes, parser state, event order, and all callers remain unchanged. For
ordinary unchanged UTF-8 attribute text, quick-xml can return `Cow::Borrowed`;
entity decoding or XML whitespace normalization can return `Cow::Owned` as
before. The candidate removes the forced `into_owned()` for event-local
`r`, `s`, `cm`, and `vm` values only. It does not claim to remove the `t`
materialization or every allocation in worksheet parsing.

The borrow is valid only during the current start-element event. The
attribute descriptor is consumed, but the borrowed arm refers to the event's
backing bytes; `start_cell` parses or converts it before the reader advances.
No `Cow` is stored in parser state or in the source snapshot. This preserves
the immutable source boundary and avoids any lifetime extension or unsafe
aliasing.

## Behavior obligations

The candidate is acceptable only if all of these remain exact:

1. `element.attributes().with_checks(true)` still examines the complete list.
   Duplicate unqualified and unknown attributes continue to be rejected by
   the checked iterator. Prefixed attributes and unknown unqualified fields
   continue to be ignored after that check.
2. Unqualified `r` is decoded when encountered. Its XML decoding error still
   precedes coordinate parsing and later fixed-field checks. Coordinate
   parsing, row matching, inferred-column bounds, and their error text remain
   unchanged.
3. Style, `cm`, and `vm` are decoded only at their existing later checks, in
   that order, with the same numeric parse and Office bounds. Their normalized
   lexical values and invalid-value messages must remain identical.
4. `t` is decoded after the coordinate/style/metadata checks and is owned
   before any later value, formula, or inline event. Shared-string lookup,
   formula expansion and cached values, inline rich text, date/scalar
   validation, and unknown type preservation must remain unchanged.
5. Entity and predefined-reference decoding, XML 1.0 normalization, malformed
   input errors, row mismatch, duplicate precedence, and prefixed-field
   behavior remain unchanged. A borrowed and an owned `Cow` must format the
   same decoded content.
6. The full worksheet validator, x14ac/MCE handling, resource reservations and
   limits, execution/cancellation fences, source bytes, relationships,
   source-version lineage, and publication/readback behavior remain intact.

This is compatible with the accepted ADR constraints because it is private to
the semantic XLSX parser, changes no public API or dependency, does not add
unsafe code, and retains all validation and source ownership. In particular,
the immutable snapshot/source requirements and the requirement to measure
copied bytes and allocations still apply; a cheaper temporary does not permit
an approximation or a validation bypass. The [0530 planning source review](../change-0530/planning-source-review.md)
and [0531 rejection review](../change-0531/rejection-review.md) document the
already-rejected broader parser/MCE directions.

## Existing coverage and smallest guard

The raw worksheet tests already cover the relevant contract surface:

* sparse numeric, explicit-empty, inline, shared-string, scalar-formula and
  shared-formula materialization;
* entity decoding in `r` and `t`, plus ordinary `s`, `cm`, and `vm`
  ([attribute test](../../../../crates/litchi-xlsx/src/raw/worksheet/tests.rs#L513));
* duplicate checks for `r`, `s`, `cm`, `vm`, `t`, unknown fields, and prefixed
  fields ([duplicate test](../../../../crates/litchi-xlsx/src/raw/worksheet/tests.rs#L527));
* coordinate, style, type, metadata-bound, malformed-entity and duplicate
  error precedence ([precedence test](../../../../crates/litchi-xlsx/src/raw/worksheet/tests.rs#L559)); and
* prefixed-field filtering ([filter test](../../../../crates/litchi-xlsx/src/raw/worksheet/tests.rs#L631)),
  plus the broader formula, unknown-type, MCE, limits and source-backed suites.

Before measuring, the smallest useful differential guard is to extend the
entity-decoding case with valid entity-encoded `s`, `cm`, and `vm` values and
assert the resulting fixed fields where the crate's test access permits. Keep
an ordinary unescaped case in the same guard (the existing plain fields cover
that borrowed path). This exercises both quick-xml `Owned` and `Borrowed`
arms without coupling a behavioral test to the private `Cow` variant. Keep
the existing duplicate and precedence tests unchanged; add only an escaped
malformed numeric case if the candidate implementation changes that order.
No allocator-failure simulation or timing assertion belongs in this test.

The differential run must compare exact stores, typed errors and precedence,
and retained source bytes through the existing raw and source-backed suites.
The type conversion at `PendingCell` is the specific lifetime guard: any
implementation that leaves `t` borrowed or retains a borrow after the event
must be rejected.

## Measurement prerequisite and disposition

The current performance harness has a material measurement gap: the planning
interval around `edit_sheets` (approximately `tools/perf-baseline/src/lib.rs`
41893--41900) has no allocator region; the next allocator region begins in
commit (approximately 41908). Commit/publication allocation vectors therefore
cannot prove a planning allocation effect for this candidate. Before a
baseline or candidate allocation claim, add a planning-local allocator bracket
around the exact `editor.edit_sheets(selectors)` call, excluding selector
construction, setup, fixture/oracle work, commit, publication, and readback.
It must report calls, bytes, reallocations, and incremental peak with the same
resource context and corpus.

After that enabler and the focused guard pass, use fresh matched baseline and
candidate builds for the prescribed native, planning-profile,
planning-allocation, commit/publication, eager-read, and quality lanes. Freeze
the thresholds and retain the complete error/store/source differential. A
lower Callgrind edge or temporary allocation count alone does not authorize
retention; the candidate must satisfy the existing end-to-end gates without
regressing any guard. No 0537 performance claim is available yet.

This candidate is distinct from the rejected 0522 lossless-layout cell
scanner, 0527 row-arena layout, 0529 attribute probe, and 0531 MCE
`memmem`/namespace-search replacement. Those directions remain out of scope;
this review does not assert that every other borrowing, visited-state, or
duplicate-access design has previously been rejected. The active priority is
OLE2/OOXML, with ODF deferred until that goal is complete and iWork excluded.
