# Change 0329: typed XLSB text-formula dependency errors

## Decision

The XLSB text compiler now distinguishes missing or ambiguous workbook
dependencies from malformed formula input. Dependency failures are reported
as the typed `UnresolvedDependency` error and are not identified by matching
error-message suffixes. This keeps context resolution explicit while
preserving the source formula for a later retry.

The compiler uses `UnresolvedDependency` for missing or ambiguous sheet, name,
table, table-column, and external metadata. Syntax is validated before
context resolution, so malformed syntax cannot be reclassified as a missing
dependency. Malformed geometry, signed metadata, and ranges remain typed
structural errors. Unsupported functions, valid external cell references, and
determinable DDE or OLE sources remain typed `UnsupportedFeature` errors.
AddIn metadata without qualifier identity is ignored so it cannot contaminate
an unrelated external-workbook lookup.

## Writer and authoring contract

The writer no longer infers deferral from message text. Typed deferral is
permitted only for context-free array and shared-formula authoring, where the
formula can be retained without pretending that workbook-dependent text has
been resolved. A full-context save propagates unresolved and structural
errors transactionally; it does not commit a partially resolved package.

Malformed array ranges fail as structural errors rather than being deferred
or silently normalized. If a save fails because a dependency is unavailable,
the caller can add that dependency and retry. The retry continues to use the
preserved source formula, rather than a lossy reconstruction from the failed
save attempt.

## Preservation and error boundaries

`UnresolvedDependency` means that the formula's syntax is understood enough
to identify the required context, but the required workbook metadata is
missing or ambiguous. It does not mean the formula was evaluated or
recalculated. Structural errors continue to reject the operation, while
unsupported features remain explicitly unsupported; neither category is
converted into an opaque success merely to make a save proceed.

The change does not add external-target loading. Authored formula text and
other preserved formula state remain available for the supported retry and
lossless-preservation paths.

## Scope and non-claims

This change covers typed dependency classification in the XLSB text compiler,
the writer's context-free array/shared authoring deferral, and transactional
full-context save behavior. It does not claim complete table or external
writer-context support. It does not fetch or load external targets, evaluate
or recalculate formulas, or claim a performance, RSS, memory, or OOM
improvement.

## Validation evidence

- XLSB library tests: 542 passed.
- XLSB integration tests: 115 passed. The exact persistent test
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` was
  skipped because it remains an unrelated observed 5-versus-6 anchor failure.
- Strict Clippy with `-D warnings` passed.
- Rustdoc with `-D warnings` passed.
- The minimal facade `xlsb` feature check passed.
- Formatting and diff checks passed.

All Cargo validation commands ran serially with `CARGO_BUILD_JOBS=1` in the
single isolated target `/dev/shm/litchi-0329-target`. No parallel Cargo
rebuild was used for this change.
