# XLS `CellParsedFormula.rgcb` compatibility design review

## Scope and decision

This is a read-only design audit for BIFF8 `Formula` records whose `cce` ends
before the record payload. The bytes after `rgce` are the formula's `RgbExtra`
(`rgcb`), rather than an extension of the token count. The review covers the
reader, semantic cell handoff, writer, and cell-values structural-edit paths.
It was performed against `docs/GOAL.md`, `docs/adr/README.md`, and the
accepted ADR set before any edit. No production change, build, test, or
profile was run for this review.

The smallest safe implementation is a shared, private `RgbExtra` framing
owner in `formula_metadata`, with a source-bound copy of the original row,
column, token stream, and exact extra bytes. The parser may retain bytes only
after it has validated their required structure and exact consumption. The
writer may emit them only when the cell anchor and token stream it is about to
emit are byte-identical to the source-bound anchor and tokens. This preserves
a valid source without attaching ancillary data to a newly encoded formula.

The fixture census currently contains 1,416 Formula records and five records
with a 10-byte suffix. Each of those five token streams begins with `0x26` or
`0x46`, consistent with a classified `PtgMemArea`; ten bytes is the size of a
count of one plus one eight-byte `Ref8U`. Length alone is not proof, so the
implementation must validate the count, every `Ref8U`, and exact end of
`RgbExtra`.

## Format facts

The local `[MS-XLS]` specification gives the relevant contract in sections
2.5.198.3 (`CellParsedFormula`), 2.5.198.59 (`PtgExtraArray`), 2.5.198.61
(`PtgExtraMem`), 2.5.198.70 (`PtgMemArea`), and 2.5.198.103 (`RgbExtra`). The
public copy of the last section is [MS-XLS 2.5.198.103 RgbExtra](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/70f743b2-a853-4c57-88be-8af637ac6e43).

`CellParsedFormula` stores a `u16 cce`, then `rgce` of exactly `cce` bytes,
then `rgcb`. `RgbExtra` is an ordered sequence tied to the Ptgs in `rgce`.
`PtgArray` requires `PtgExtraArray`, `PtgMemArea` requires `PtgExtraMem`, and
the three `PtgElf*` forms require `PtgExtraElf`. Name, external-name, and 3-D
references require revision structures only when the containing formula is in
the corresponding revision context. Consequently, arbitrary bytes after
`rgce` cannot be accepted as opaque padding.

The initial bounded compatibility scope should implement `PtgExtraMem` and
`PtgExtraArray`, which are the ordinary structures needed by the current
fixtures. `PtgExtraElf` and revision-only structures need a separately checked
context and should produce an explicit unsupported/refusal result until that
context exists. A future extension can add their parsers without changing the
source-bound ownership rule.

## Current data path and gaps

* `crates/litchi-xls/src/formula_metadata/codec.rs:35-89` requires
  `FORMULA_FIXED_SIZE + cce == data.len()`. It therefore rejects every
  `CellParsedFormula` with `RgbExtra` before `CellRecord` is constructed.
  `Parsed` currently owns only `formula: Vec<u8>` (the `rgce` bytes).
* `crates/litchi-xls/src/records.rs:1319-1363,1620-1647` stores formula
  metadata and `rgce` in `CellRecord::Formula`; adding a second suffix field
  would duplicate ownership. The recommended design puts the validated suffix
  in `Metadata`, so the existing record variant remains source-compatible.
* `crates/litchi-xls/src/cell.rs:62-195,218-228` already clones formula
  metadata into `Cell` and exposes `formula_metadata()`. The suffix therefore
  reaches semantic cells without a new `Cell` field. `formula_bytes()` remains
  the token stream; a new metadata getter is optional and should expose only a
  validated owner, not an unvalidated byte slice.
* `crates/litchi-xls/src/formula_metadata/model.rs:57-210` is the natural
  owner. `Metadata` already carries source-derived flags, cache, shared owner,
  and array owner. Its new optional ancillary value should be private and
  constructible only by the checked parser (or by a checked crate-internal
  constructor). `Metadata::new`, `from_wire`, and ordinary authoring builders
  must continue to create an empty ancillary value.
* `crates/litchi-xls/src/formula_metadata/array/codec.rs:29-62` already splits
  `cce`, `rgce`, and the remaining bytes, retains the exact `rgcb`, and calls
  strict validation. `array/validation.rs:39-47,798-875` demonstrates the
  required exact-consumption and bounded `SerAr` checks. It is useful as a
  validation model, but its owner is for the separate `Array` record and must
  not be attached directly to a cell formula.
* `crates/litchi-xls/src/list_object/codec/binary/formulas.rs:8-142` has the
  only existing token-to-extra scanner. It recognizes classified `PtgArray`
  and `PtgMemArea`, checks token lengths, and bounds `PtgExtraMem` and
  `PtgExtraArray`. It is private to `list_object`, returns only an end offset,
  and reports list-specific errors. It should be factored into a private
  `formula_metadata` grammar primitive rather than imported by another crate
  or copied wholesale. The list-object adapter should retain its current
  stricter list-formula token policy and error behavior.
* `crates/litchi-xls/src/writer/biff/cells.rs:200-300` chooses the actual
  token stream after shared/array owner substitution, writes `cce`, and emits
  only `22 + cce` payload bytes. The ancillary check must happen after that
  choice and before `write_record_header`; the suffix length must be included
  in the payload limit, while `cce` remains the token length.
* `crates/litchi-xls/src/cell_values/mod.rs:5885-5929` indexes a Formula's
  cached value but does not validate `cce` against the rest of the record.
  `:6756-6801` changes only the eight-byte cached value in place, so it already
  preserves a valid suffix exactly. `:4123-4155` fingerprints from the flags
  through the end of the record, and therefore already includes `cce`, `rgce`,
  and `rgcb` in the dependency precondition.
* `crates/litchi-xls/src/cell_values/structural.rs:2386-2419` currently
  insists that `22 + cce == payload.len()`. It must use a common framing helper
  and return a typed `UnsafeEdit` for a valid ancillary formula before any
  row/column token patch. A malformed suffix remains a source validation
  error. The movement code cannot safely shift every `Ref8U`, array value, and
  revision reference, so it should refuse every nonempty `RgbExtra` until a
  complete remapper exists.
* `cell_values/mod.rs:4157-4181` compares a newly authored canonical record
  against a source record without a source-bound ancillary representation. The
  resource schema has no suffix field, so matching must reject a target with a
  nonempty ancillary suffix rather than claim that a canonical no-suffix
  resource operation preserves it. The smaller safe change is the typed
  refusal; an exact resource operation can later carry the source-bound owner.

## Proposed common owner and parser

Add a private value conceptually equivalent to:

```text
Ancillary {
    original_cell: (u16, u16),
    token_len: usize,
    source: Vec<u8>, // original tokens followed by exact RgbExtra bytes
}
```

and add `Option<Arc<Ancillary>>` to `formula_metadata::Metadata`. The single
bounded source buffer avoids duplicating the token and suffix allocations; the
outer `Arc` shares the checked owner through snapshots. The owner is
source-bound deliberately: a `Vec<u8>` suffix alone cannot prove that it still
belongs to the token stream or cell coordinates a later writer supplies.
`PtgExtraMem` contains coordinate-sensitive `Ref8U` data, so matching tokens
while moving the same formula to another row or column is insufficient. `Arc`
shares the checked source data through `CellRecord`, `Cell`, and metadata
snapshots without introducing a public raw-storage type. The constructor must
be crate-private, must receive the checked source cell, already validated
tokens, and bytes, and must not accept an arbitrary caller suffix.

The shared parser should have two bounded layers:

1. A token-framing pass recognizes the token sizes and records required extra
   kinds in order. It should reuse the existing classified-opcode and
   variable-length rules, but not import list-object's formula-authoring
   restrictions into ordinary formulas. Existing formulas with no `RgbExtra`
   and an opaque token sequence must retain their current acceptance. If a
   nonempty suffix is present and the token stream cannot be framed
   unambiguously, fail closed rather than guessing its owner.
2. An extra pass reads exactly the required structures in that order. For
   `PtgExtraMem`, check the bounded `u16` count, parse every eight-byte
   `Ref8U` with row/column ordering and the `0x00ff` column ceiling, and ensure
   the final offset is exactly `bytes.len()`; this grammar has no additional
   reserved-bit check for the full-width `Ref8U` coordinates. For
   `PtgExtraArray`, check dimensions with checked arithmetic and reuse the
   strict bounded `SerAr` value parser, including UTF-16 and error-code checks.
   Reject unowned trailing bytes, truncation, integer overflow, and any
   allocation above the configured formula/extra limits. For a nonempty suffix,
   a required extra with no bytes is malformed and bytes with no required extra
   are also malformed. An empty suffix remains the legacy opaque/no-extra path.

`parse_record_with` should compute `token_end = 22 + cce` and require
`token_end <= data.len()`, not equality. It should validate the trailing slice
through the common parser, retain it unchanged in `Metadata`, and keep the
existing empty-token rule for non-string cached values. `parse_record` and
`parse_record_preserving` must take the same path so the compatibility profile
cannot bypass ancillary validation. The ordinary cell path has no revision
context: Name/NameX/3-D tokens are framed at their ordinary widths, while any
revision-owned tail fails exact consumption. A nonempty ELF tail remains an
explicit typed unsupported result; it must not be stored as unchecked bytes.

`formula_metadata::validate_for_write` should recheck the ancillary owner and
its source token relationship. In `write_formula_with_metadata`, first select
the actual emitted token stream (including shared/array owner behavior), then:

* if there is no ancillary owner, preserve the current output exactly;
* if an owner exists and either the output cell differs from
  `original_cell` or `actual_tokens != original_tokens`, return a typed
  stale/source-bound refusal before writing any record bytes;
* otherwise add `bytes.len()` to the `22 + tokens.len()` size calculation,
  enforce both the configured bound and the BIFF8 8,224-byte payload limit,
  write the original `cce`, tokens, and exact `bytes` in that order.

This check also prevents attaching a cell's suffix after shared or array owner
substitution has changed its tokens. Any owner transition, token edit, or
canonical re-encoding must either carry a newly validated ancillary owner or
fail before output. New authored formulas continue to have no ancillary bytes
and retain the existing canonical output.

## Structural and preservation policy

Valid `RgbExtra` is retained for read-only access, exact no-op serialization,
style changes, and cache changes. The existing in-place cache writer already
has the required byte-preserving behavior. Semantic worksheet construction
must not discard the metadata when it binds shared or array owners; the
existing `CellRecord` to `Cell` metadata clone is sufficient, but the owner
binding code in `workbook/codec/semantic/worksheet.rs` needs a regression check
that it does not replace the ancillary field.

Row and column insertion/deletion must refuse any Formula record with a
nonempty, valid `RgbExtra` using `Error::UnsafeEdit`. This is intentionally
broader than checking only `PtgExtraMem`: array and revision ancillary values
can carry references or formula-dependent structures, and silently moving
only `rgce` would break dependency closure. The alternative is a complete
ordered remapper for every supported ancillary structure, including all
`Ref8U` locations and revision-owned references; that is a separate feature,
not a compatibility fix.

The structural parser should still validate Formula framing while indexing.
The movement helper can return a small internal result containing tokens and an
`has_ancillary` bit, allowing a recognized valid suffix to become
`UnsafeEdit` and preserving `InvalidData` for genuinely malformed source.
`patch_formula_record` must inspect this result before mutating its candidate.
`formula_dependency_fingerprint` should continue hashing through record end.
Formula resource insertion/removal should refuse an existing ancillary target
until the resource format can carry the source-bound bytes. The matching path
must inspect the complete BIFF record, including its payload length and
`RgbExtra`; its length-aware comparison must not claim that a canonical
no-suffix formula matches a source record with an ancillary tail.

## Ownership and implementation file map

The production change should be coordinated as one XLS ownership change:

* `formula_metadata/mod.rs`, `model.rs`, `codec.rs`, and `validation.rs` own
  the ancillary type, parser, limits, and write validation.
* `list_object/codec/binary/formulas.rs` adapts to the extracted private
  token/extra primitive and keeps its existing list-specific entry points.
* `records.rs` needs only the parser call if metadata-only ownership is used;
  no second `CellRecord::Formula` suffix field is recommended.
* `writer/biff/cells.rs` appends the validated source-bound bytes and checks
  actual-token identity before emitting.
* `cell_values/mod.rs` validates Formula framing during indexing, leaves
  cache-only writes in place, and refuses canonical resource matching when a
  suffix is present.
* `cell_values/structural.rs` uses the common framing result and refuses
  coordinate movement with ancillary data before patching.
* `cell.rs` and `workbook/codec/semantic/worksheet.rs` should need no layout
  field; inspect their metadata cloning/owner-binding paths and add only a
  getter or preservation assertion if the implementation requires it.

## Focused verification plan

These cases should be added to the existing XLS tests after implementation;
they were not run as part of this review:

1. Parse a Formula with one `PtgMemArea` and a valid 10-byte `PtgExtraMem`,
   including both classified token forms in the fixture census. Check the
   retained bytes, semantic cell metadata, and exact no-op writer output.
2. Parse valid `PtgArray`/`PtgExtraArray` and a mixed ordered extra sequence;
   check dimensions, `SerAr` bounds, and exact bytes. Confirm a no-extra formula
   remains byte-for-byte compatible with current output.
3. Reject no-required-extra plus trailing bytes, a truncated count or
   `Ref8U`, count/structure mismatch, arithmetic overflow, invalid range
   fields, malformed `SerAr`, trailing bytes, and unsupported revision/Elf
   structures. Confirm no suffix is retained after any failed validation.
4. Attempt row and column movement on each valid ancillary form and assert a
   typed `UnsafeEdit` with the original stream unchanged. Exercise the
   ordinary Ref/Area path to prove its behavior is unchanged.
5. Change only a formula cache or style and assert that the ancillary suffix
   remains byte-identical. Exercise string-pending formulas and their
   intervening companion records.
6. Rewrite a parsed formula with changed tokens, shared-owner substitution,
   or array-owner substitution and assert that the writer refuses the stale
   source-bound owner before emitting bytes. A newly authored formula with no
   suffix must retain the existing canonical writer behavior.
7. Keep the existing list-object formula tests and verify extraction preserves
   their token legality, error ordering, and `PtgExtraArray`/`PtgExtraMem`
   framing behavior.

## Risks and unresolved scope

The specification permits more `RgbExtra` forms than the five current native
records exercise. Treating every unknown tail as opaque would violate the
ordered-correspondence rule, while attempting a full revision grammar in this
compatibility patch would enlarge the ownership and test surface. The bounded
Mem/Array scope with ordinary Name/NameX/3-D token framing and explicit typed
refusal for ELF/revision-owned tails is therefore the safe first increment.
The fixture census should be rechecked after parser work, especially for count
values and range limits. No implementation may silently drop a validated
suffix or append one after re-encoding a different token stream.
