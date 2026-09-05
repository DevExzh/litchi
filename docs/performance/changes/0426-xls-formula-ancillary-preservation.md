# 0426: XLS Formula ancillary preservation and facade correctness

The native `FormulaEvalTestData.xls` fixture contains five standard
`CellParsedFormula.rgcb` tails among 1,416 Formula records. The previous reader
required `22 + cce` to equal the entire record length, rejecting those valid
records. The XLS codec now separates token bytes from bounded ancillary data,
validates ordered `PtgExtraMem` and `PtgExtraArray` structures, and retains exact
bytes in shared formula metadata. It checks range ordering and column limits,
memory-expression boundaries, strict array scalar values, exact consumption,
and the 8,224-byte BIFF record limit.

The writer requires the original cell coordinates and the actual emitted token
stream before emitting any record bytes. Changed tokens, shared/array token
substitution and coordinate movement refuse with `UnsafeEdit`. Cache/style
edits preserve the tail; canonical Formula resources refuse ancillary records
they cannot represent. Empty tails retain the existing opaque-token acceptance
contract. This is bounded ancillary support, with no new complete formula/RPN
validator, evaluator, missing-extra certification, ELF/revision support or
structural ancillary remapper.

Five other baseline-reproduced facade failures require fixture/assertion
corrections: valid ODT content owns malformed inert OOXML-looking extras and
misleading filename suffixes, deferred XLSX access needs payload corruption
that leaves ZIP framing intact, and wrong-format XLSB tests need a valid OPC
root relationship. Valid OOXML/ODF polyglot limit checks and typed wrong-format
refusals remain covered. The [bundle](../results/change-0426/README.md) retains
the independent CFB census, source-only reviews, all failed attempts, exact
commands and final verification outcomes.

The final XLS suite passes 1,341 tests with one ignored doctest; all 461 facade
tests pass with 11 ignored. The 12 new public integration cases are included
in the XLS total, not added again. All eight dependent XLS/CFB harness tests
pass. Strict XLS Clippy and warning-denied XLS/facade documentation pass. The
facade strict rerun retains its 18 prior findings; ODF layout, standalone
harness lint and native-resave lockfile debt from 0425 are still open.

This is correctness work with no performance claim. The new optional metadata
owner has a static layout cost even for formulas without a tail; the bundle
records type sizes separately from operation allocations, peak live memory and
RSS: metadata grows 24→32 B, CellRecord 80→88 B and semantic Cell 152→160 B.
Nonempty tails use one combined token/tail buffer shared through an outer
`Arc`; empty tails add no ancillary payload allocation. No matched timing,
memory or I/O reduction is inferred.

The full non-iWork performance program remains active. Global strict-gate debt,
caller-drop and near-limit memory evidence, native producer matrices, cold and
range sources, bounded semantic streaming/append, scaling and broader CRUD
coverage still require work.
