# Execution notes

Production was never edited. Candidate source was archived before root ran any
build or capture. Source review caught and fixed an early draft transition that
could fall from replay into the map fallback; the measured candidate re-enters
the checked path at count one.

The initial helper quality attempt passed baseline tests and Clippy, then
failed candidate test compilation with E0631: `Iterator::any(Result::is_err)`
passes an owned item to a method expecting a reference. The fix is exactly
`any(|item| item.is_err())`; production candidate logic did not change.
`quality-failed-0` retains logs and receipts; `test-src-failed-0` retains the
exact failed mirror sources and locks. The relocation witness and independent
failure audit preserve the original identities. The final helper attempt passed
all six rows, including 65 baseline and 85 candidate test executions with zero
failures or ignored tests (13 and 17 tests repeated over five helper copies).

The direct probe's locked build/check/Clippy, formatting, catalog, and semantic
self-check all passed before native timing. Input literals, both exact helper
modules, the manifest and dependency lock, and the preflight policy are frozen
by build receipts. The separate minimal mirror crates do not certify full
production-crate integration or public-workflow behavior.
