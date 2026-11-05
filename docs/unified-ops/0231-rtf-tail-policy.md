# 0231 — Checked RTF tail policy

`tail_append::SemanticLimits` replaces the public-field `TailAppendLimits`
constructor. The consuming builder checks each bound against `isize::MAX` and
returns `Error::InvalidLimit { kind, value, maximum }` before an edit is created.
Zero explicitly denies the selected resource. `into_builder` preserves existing
settings when adjusting a bound; callers can no longer mutate profile fields.

Production defaults remain 256 paragraphs, 4,096 runs, 256 KiB UTF-8 input,
512 KiB encoded insertion, 64 MiB output and 2 MiB durable envelope. Existing
staging, output and durable decoding paths use the admitted profile. These
scoped bounds do not replace retained source policy or execution budgets.

`EditOptions` carries the profile through
`Package::edit_tail_append_with_options` and `Edit::new_with_options`.
Source-backed edits use `TailAppendSourceEdit::new_with_options`. Both edit
forms expose `options()`; unchecked `with_limits` constructors and the old
package entry are removed. Default entries remain available. Native tests and
the benchmark harness construct checked profiles and consume typed errors.

## Verification

The full native suite passes 1,462 tests/doctests. Three new policy tests cover
all six defaults, zero and maximum construction, typed overflow errors,
profile-preserving builder conversion, per-edit option retention, zero-count
refusal, failed-admission reuse and source immutability. Existing retained and
source-backed tail publication and durable replay tests pass after migration.

Warning-denied native all-target Clippy and rustdoc pass. The glossary checks
85 banned names and the six-member tail family; 21 Python checker tests pass.
The benchmark library compiles with the migrated calls and facade dependencies.
The full facade test suite was not rerun for this scoped native policy change.

Validation uses external copied locks/build output. All three repository
lockfiles remain unchanged. Owned 0231 temporary files are removed after
qualification; no worktree was created and no large evidence is committed.

## Remaining work

Publication-window limits still need checked profiles and options; source and
publication family names, redaction vocabulary and unified RTF CRUD remain
open. Other formats, Python/CLI, corpus and performance acceptance remain
required by the full active goal. No runtime performance acceptance claim is
made in this wave.
