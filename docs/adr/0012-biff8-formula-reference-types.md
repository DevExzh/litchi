# ADR 0012: Checked BIFF8 formula references and panic-free encoding

- Status: Accepted
- Date: 2026-08-03

## Context

The public XLS formula tokenizer accumulated A1 column letters in a `u16`
without checked arithmetic. Long inputs could therefore panic when overflow
checks were enabled and wrap when they were not. It also accepted columns past
BIFF8's `IV` limit and the public `PtgRef` and `PtgArea` tuple variants let a
caller construct the same invalid column values directly. Encoding those
values copied reserved bits into BIFF8 column fields.

This violates the panic-free facade rule and makes format validity depend on a
convention that the public type system does not enforce.

## Decision

The focused formula module owns short checked reference values. `formula::Ref`
stores a BIFF8 row, column, and relative flags; `formula::Area` stores two
ordered endpoints. Their fields are private. A zero-based BIFF8 row is a
`u16`, which exactly covers all 65,536 rows, while a zero-based column is a
`u8`, which exactly covers `A` through `IV`. Public constructors and accessors
remain concise, and no unchecked constructor is provided.

`Ptg` carries these values through contextual `Ref`, `Area`, and `Area3d`
variants instead of exposing raw coordinate tuples. The 3D form adds only its
checked sheet index to an `Area`. Because every reference-bearing token is
valid by construction, binary encoding remains infallible for coordinate
bounds and cannot emit reserved column bits.

A1 parsing uses checked arithmetic and rejects a column as soon as it exceeds
the BIFF8 grid. Invalid, oversized, or otherwise adversarial formula text
returns the crate's typed error and never unwinds. Named ranges, conditional
formats, data validation, and workbook formula records reuse the same tokens;
they do not reintroduce raw columns at an adapter boundary.

The cell-level `Formula` owner also types the metadata surrounding the token
stream described by [MS-XLS] 2.4.127: `fAlwaysCalc`, `fFill`, `fShrFmla`,
`fClearErrors`, and the opaque `chn` calculation cache. Reserved flag bits,
zero-length `cce`, truncated payloads, and an orphaned shared-formula flag are
rejected at the BIFF boundary. `Cell::formula_metadata` exposes the decoded
value, while the writer accepts non-shared metadata and refuses to emit
`fShrFmla` until a `ShrFmla` sequence owner exists.

The refactor is intentionally breaking. Redundant `Ptg` prefixes and public
tuple construction are not retained as aliases. Context supplies the missing
meaning, for example `Ptg::Ref(Ref::new(...))`.

## Consequences

- Any public token accepted by the BIFF8 encoder has a representable row and
  column by construction.
- Text formulas, defined names, conditional formats, and validation formulas
  share one grid boundary and one failure policy.
- Formula metadata round-trips independently of formula evaluation, preserving
  BIFF8 calculation flags and the opaque cache without claiming recalculation.
- The safe path does not need a defensive runtime branch for a column value
  that its input type cannot hold.
- This safety correction makes no parser-throughput or allocation claim.

## Verification

Focused tests cover `IV` as the last accepted column; `IW`, `ZZZZ`, overflow-
length column strings, and malformed rows as typed rejections; direct checked
reference and area construction; exact encoded flag bits; Formula metadata
decoding and validation; and public writer serialization returning an error
without unwinding. The formula-metadata and focused writer regressions pass;
the previously green broader formula gates remain the reference verification
for this ADR. Per explicit user direction, the full-workspace test suite is
not repeated.
