# 0362ul66 — bounded Numbers-wire raw operations

Status: source READY for independent review.

## Scope

The Numbers-wire leaf previously let callers parse and rewrite BNC values and
table payloads without carrying the owner-selected policy through the raw
boundary. Owned parsing also copied every field and opaque tail without a
preflight. This wave makes the raw API explicit and bounded across Numbers,
Pages and Keynote.

## Canonical policy

`litchi_numbers_wire::Limits` is the one raw policy value. Its typed
`LimitKind` dimensions cover input bytes, output bytes, fields, opaque tail
bytes and repeated items. Invalid profile values and observed overages are
typed errors; no operation silently resets or clamps a caller's selected
profile. `Limits::from_wire_limits` carries the common IWA wire bounds, while
the Numbers-wire profile adds tail and item ceilings.

These public boundaries now require `&Limits`:

- `BncCell::{parse, encode}` and `BncCellView::parse`;
- `pre_bnc::PreBncCellView::parse` and `cell_value::decode_cell_value`;
- `table_cells::read_table_cells` (its existing aggregate read budget is
  retained);
- `table_write::{tile_cell, tile_set_cell, header_adjust, list_intern,
  list_release, cell_bytes, cell_formula_state, table_formula_state}`.

Readers preflight input, fields, tails and items before retaining data.
Variable writers preflight exact output lengths and source cell lengths before
fallible `Vec` reservation; allocation failures are represented as typed
errors. Owned BNC fields and tails retain their source bytes and re-encoding is
verified against the original wire payload. Unknown protobuf fields and BNC
tails continue to be preserved by the borrowed and owned paths.

## Consumer migration

The Numbers, Pages and Keynote owners derive a raw profile from their retained
archive/source and semantic ingress limits, then pass that profile through all
raw parse, decode, encode and table-writer calls. Tests, fuzz targets and
fixtures use an explicit `Limits` value; no default-only compatibility
overload or alias remains. The API glossary registers the `numbers-wire`
profile and required signatures for the gate.

## Validation

The source-only checks cover exact and one-under input, field, tail and output
limits, plus unknown-tail owned round-trip fidelity and table writer policy
rejection. Cargo, native and rustdoc validation are intentionally deferred to
the coordinator's heavy-build slot for this wave; no performance qualification
claim is made here.
