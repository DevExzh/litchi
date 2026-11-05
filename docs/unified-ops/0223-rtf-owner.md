# 0223 — Immutable RTF package and semantic projection

The native immutable RTF owner is now the concrete `Package`. Parsing, opening,
editing, replay and binary publication belong to that owner. `Package::document`
returns an owned read-only `Document` sharing the same immutable model and lazy
text cache. The projection survives the original handle, retains its old
snapshot after an edit, and exposes no open/edit/binary-publication methods or
`Deref` escape to the owner. Plain-text export is a semantic read operation.

The old `edit::Snapshot` alias is removed. Existing editor, history, durable
patch, redaction, append and writer consumers use `Package` directly. Advanced
read-only parser-model access is explicit through `Package::raw()`. Retained
source limits and open policy are inspectable on the package.

The unified RTF handle now retains the native immutable package. Its existing
family views borrow the package's advanced model; this preserves their current
read behavior while establishing the owner needed by the future CRUD adapter.
External resave and benchmark consumers migrate with the leaf API.

The corpus caught a real distinction between retained semantic body paragraphs
and the facade's existing raw paragraph projection (`chtoutline.rtf`: six versus
five). This migration preserves the facade's existing raw count to keep it
consistent with its existing paragraph accessors. The future semantic adapter
must migrate count and traversal together; this is not claimed as completed.

## Verification

Native internal migration: `c3875c6cff`; consumer migration: `1e2fdc6dfc`.
Validation passes:

- `cargo test -p litchi-rtf`: 1,427 tests/doctests, including three projection
  lifecycle compile-fail checks and shared-cache/lifetime assertions.
- `cargo test -p litchi-ops --all-features`: 1,610 tests/doctests.
- RTF-only facade policy suite: five tests.
- All-target warning-denied Clippy and warning-denied rustdoc for both crates.
- Rustdoc JSON glossary: 37 bans; checker suite: 21 tests.
- Benchmark library and RTF native-resave consumer compile checks.
- Crate boundaries: 66 packages; OfficeCLI delta and its 15 Python tests.

Validation used external copies of all three Cargo locks. Root validation used
cached cc 1.6.0; benchmark/resave checks retained their lock contents. Repository
locks are unchanged. Base `32ad64e320` remains an ancestor. Temporary 0223
worktrees, branches, targets, copied locks and logs are removed after validation.
No compiled artifacts or large evidence are committed.

## Remaining work

The primary transaction vocabulary still needs canonical commit/patch methods
and finite semantic profiles. Canonical Builder/minimal creation, atomic native
save, unified RTF CRUD and full coverage remain required. This owner migration
does not claim those contracts or performance acceptance. The full goal stays
active, including other tier-3 formats, Python/CLI completion and corpus gates.
