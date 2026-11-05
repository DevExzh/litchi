# 0032 — Typed primary DOCX transaction limits

`document::Error::Limit` now carries `kind: document::LimitKind`, `observed`
and `maximum`. The free-form `resource`, `actual` and `max` fields are removed
without aliases. The non-exhaustive, copyable kind enum distinguishes the
existing XML, staging, transfer, durable/history and managed-accounting bounds.
Display keeps structural resource labels; Debug exposes typed kinds and counts,
never authored content.

Every primary transaction construction site is migrated, including focused
package transfer/story adapters and the unified operations error mapping test.
Dynamic durable count/path helpers take a typed kind directly. Existing numeric
ceilings and refusal order are preserved, except for the aggregate transfer-graph
correction below. Arithmetic representability failures retain the established
saturated `usize::MAX` observation. When both counts are `usize::MAX`, equality
means unrepresentable size, not an admitted value; the public field docs explain
this sentinel. The character-cursor ceiling is conservatively derived from its
UTF-8 byte ceiling and its kind documents that relationship.

Review found the durable transfer-graph aggregate ceiling still returned
`InvalidDurable` text. It now returns `LimitKind::TransferGraphBytes` at one shared
64 MiB constant, including overflow. The decoder checks the aggregate before
copying the next payload. Count conversion failures also return the requested
typed limit. Other malformed-wire errors remain format errors.

Both DOCX rustdoc profiles ban the old field names specifically beneath
`document::Error::Limit`. That scope does not rename unrelated leaf, OPC,
allocation or source-story errors. The complete six-type family is deliberately
not registered yet: exposing the names alone does not prove the remaining owner,
policy and caller-selected-limit requirements.

## Verification

- Full-feature DOCX/operations unit and integration tests: 2,400 passed,
  one existing ignored (2,012 DOCX, 388 operations).
- Warning-denied DOCX doctests: 84 passed, 31 existing ignored.
- Warning-denied DOCX/operations library Clippy, including final rerun: passed.
- Full-feature facade and Python library compilation: passed.
- DOCX/operations tests/examples compilation and standalone performance harness
  compilation: passed.
- Glossary rustdoc: three profiles, 22 banned names, zero complete families;
  passed. All 19 glossary checker unit tests passed.
- Final public limit/Debug regression target: two passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations,
  11 existing explicit debt items; passed.
- Formatting and diff whitespace checks: passed.

New tests exercise the public depth ceiling and exact operation boundary, check
that a refused operation leaves staged XML unchanged, and assert content-free
limit diagnostics. Durable tests check aggregate boundary/overflow and a count
refusal before trying to read absent records. Existing managed/source, transfer,
composition and history suites exercise the migrated construction paths.

## Remaining work

Caller-selected `SemanticLimits`, canonical source limits/options, the broader
error taxonomy, specialized ownership and complete family registration remain
open. This record does not claim that all DOCX resource failures have migrated:
wrapped OPC/source-story errors keep their own types. Other formats, Python/CLI,
performance and corpus acceptance remain part of the original goal. No measured
performance or allocation-budget claim is made.

## Cleanup

Removed temporary SVG fixtures (the same two recorded in 0019), workspace and
standalone harness build output, generated root Cargo.lock, migration script and
check logs. Preserved tracked standalone lockfiles, base checkout, iWork files
and existing worktrees. No new worktree was created.
