# 0229 — RTF logical-tail commit and replay

Logical-tail commits now expose `package`, `into_package` and
`into_parts -> (Package, TailAppendPatch)`. The old `snapshot` accessor is
removed. Native replay belongs to `Package::apply_tail_append`; the patch's
implementation is crate-private, and `is_noop` reports exact artifact equality.
Native consumers and the live benchmark harness use the new contract.

Replay previously cloned the cached target after comparing source bytes. That
could replace an equal-byte recipient's opening policy and bypass a stricter
source-byte ceiling. Replay now returns the recipient itself for an exact
no-op, reuses the admitted target when policies match, and reopens the target
under the recipient's policy when they differ. Inverse replay follows the same
rule. Source conflicts remain typed and failed admission preserves the source.

## Verification

The full native suite passes 1,458 tests/doctests, including 23 logical-tail
tests. New regressions cover stricter-recipient rejection with a typed
`SourceLimitKind::InputBytes` refusal, successful replay and inverse under a
distinct admitted profile, no-op identity, commit decomposition, and consuming
package access. Existing exact-splice, durable inverse and publication tests
remain green.

Warning-denied native all-target Clippy and rustdoc pass. The glossary checks
54 banned names, including the removed public accessor and replay method;
21 Python glossary checker tests pass. The benchmark library compiles with the
migrated native calls, including facade dependencies. The full facade test suite
was not rerun for this native scoped-family change.

Validation uses external copied locks and build output; all three repository
lockfile hashes are unchanged. Owned 0229 temporary files are removed after
qualification. No worktree was created and no large evidence is committed.

## Remaining work

Tail-append family type names, unchecked scoped limits, source-publication
policies, and redaction vocabulary still need migration. Unified RTF CRUD and
all remaining requirements of the full goal remain active. This wave makes no
runtime performance acceptance claim.
