# DOC saved-selection reader validation

`Document::saved_selection()` exposes the inert `Selsf` record selected by
FIB `fcWss/lcbWss`. Admission retains only the exact 36-byte range; malformed
optional pointers and lengths remain deferred accessor errors and do not
prevent ordinary document opening. Successful and failed parses are cached.

The reader validates the mandatory flag, style, character-position and
block/table constraints while preserving the exact source bytes. Shape
selections ignore undefined `fInsEnd`; column selections require table state.
The document accessor checks selection positions against FIB `ccpText`.
Advisory table-column ordering is not promoted to a mandatory refusal.
This reader neither applies host UI selection state nor writes the record.

Root Rust 1.95.0 validation passed 1,201 unit/integration tests and fourteen
doctests, with two integration and twelve doctest ignores. Strict all-target,
all-feature Clippy, formatting and whitespace checks passed. Independent
review checked the local Markdown specification and final corrections.
The receipt binds nine source/documentation files and compressed/raw logs.
The ignored Cargo lock was copied from the earlier isolated reader gate and
its exact hash is recorded.

Synthetic coverage includes valid document access, shape/column flags,
character-position limits, deferred malformed lengths and cached errors.
No native-producer UI acceptance or measured performance claim is made.
