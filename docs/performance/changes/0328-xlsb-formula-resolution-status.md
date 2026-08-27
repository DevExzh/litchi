# Change 0328: explicit XLSB formula resolution status

## Decision

XLSB formula ingestion now reports whether a formula has a usable textual
representation or has been preserved opaquely. This makes fallback behavior
observable without treating an unsupported, unresolved, or unvalidated formula
as if it had been successfully interpreted. The status is owned by the XLSB
package `Cell` API and is separate from `litchi_core::CellValue`.

The public status is:

```text
FormulaResolutionStatus::Resolved
FormulaResolutionStatus::Opaque(FormulaOpacityReason)
```

The opacity reasons are:

```text
FormulaOpacityReason::Unsupported
FormulaOpacityReason::Unresolved
FormulaOpacityReason::Unvalidated
```

`Cell::formula_resolution_status()` returns `None` only for a non-formula
cell. `Resolved` means that a usable formula text representation was produced;
it does not mean that the expression was independently validated, evaluated,
recalculated, or checked for cached-value freshness.

## Preservation contract

- Binary ordinary formulas retain the exact `rgce` and `rgcb` streams through
  the parsed formula representation, including when text resolution is opaque.
- Grouped formulas retain both the exact cell placeholder and the group
  definition/range. A resolved group exposes its text while an opaque group
  still preserves the source streams and cached result.
- `Cell::cached_value()` exposes the value stored in the source record. It is
  not an evaluated or recomputed value.
- Legacy `CellRecord` construction does not parse or canonicalize its formula
  bytes. `raw_formula_bytes()` exposes those exact bytes, and such cells are
  marked `Opaque(Unvalidated)`.
- Authored formula values with non-empty text are marked `Resolved`; authored
  empty formula values and scalar values passed through the formula constructor
  are retained and marked `Opaque(Unvalidated)`.

This preserves unsupported content as readable source data while making the
loss of semantic resolution explicit. It does not silently replace an
uninterpreted formula with a guessed expression.

## Error taxonomy

`UnresolvedDependency(String)` is now a typed formula and package error for
formula text that requires workbook metadata which is missing or ambiguous.
The binary semantic reader uses it for unresolved names, XTI/sheet or
external-book references, table or table-column references, and pivot-scope,
view, or name metadata that cannot be resolved. These cases become
`Opaque(Unresolved)` at the cell boundary while preserving source bytes and
the cached value.

Known unsupported token families, including AddIn and DDE/OLE-style features,
remain typed unsupported errors and become `Opaque(Unsupported)` at the cell
boundary. Malformed token streams, invalid indices or flags, contradictory
group metadata, invalid geometry, and arithmetic overflow remain typed
structural errors; they are not converted into opaque success. In particular,
a missing or malformed group definition continues to fail construction rather
than being hidden behind a status value.

## Range safety

Formula range validation and A1 diagnostics are checked for maximum `u32`
coordinates. Out-of-range row or column values now return a typed error rather
than panicking, wrapping, or emitting misleading coordinates.

## Scope and non-claims

This change covers status reporting and lossless preservation at XLSB cell
construction, including ordinary, grouped, and legacy formula paths. It does
not add formula evaluation or recalculation, fetch or open external targets,
validate formula semantics independently, or assert cached-value freshness. It
also makes no performance, RSS, memory, or OOM claim. Writer/editor
string-matching fallback and broader text-compiler mapping remain outside this
change.

## Validation evidence

Completed validation before the final strict-gate rerun:

- XLSB library tests: 535 passed.
- XLSB integration tests: 115 passed, including all 3
  `formula_extended` tests. The exact persistent test
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` was
  skipped because it remains an unrelated observed 5-versus-6 anchor failure.

Strict Clippy with warnings denied, rustdoc with warnings denied, the facade
minimal `xlsb` feature check, `cargo fmt --check`, and `git diff --check` all
passed. Cargo commands ran serially with `CARGO_BUILD_JOBS=1` in one isolated
target. No parallel Cargo build was used for this change.

## Follow-up

The remaining formula work is to propagate the same structural-versus-
unresolved distinction through the broader text compiler and writer/editor
fallback paths, and to continue XLSB dependency parity work. Those follow-ups
must retain the preservation and typed-error boundaries described here.
