# Validation and limitations

Final harness release tests: **372 passed, zero failed, one existing ignored**.
Full OPC all-feature release tests: **458 passed, zero failed, one existing
ignored**, including source-backed topology, signature refusals, external ZIP64,
source/reader and accounting tests. These are existing fixture/test coverage;
no native Office application was launched for the synthetic OPC benchmark.
Workspace no-iWork and standalone all-feature/all-target release checks pass.
Standalone rustdoc passes with warnings denied, explicit four-file rustfmt
passes, and the crate-boundary audit passes.

Strict harness Clippy still fails on inherited debt: the retained before/after
comparison shows zero new diagnostics, including none in the new module. This
is a no-new-debt result, not a clean strict lint result. Numeric function-line
counts are normalized while diagnostic multiplicities remain checked.

All intermediate failures are retained: initial use of a nonexistent archive
reader method; the initial resource test's missing imports; one new modulo
style lint corrected before final testing; and the first Python pilot's
assumption that the generic corpus catalog enumerated members. That catalog
explicitly reports member enumeration unavailable. Complete independent member
verification instead uses the exported, hash-bound actual ZIP fixtures. The
corrected six pilots pass. No retained performance run preceded the final
protocol freeze.

The Python oracle independently regenerates deterministic payloads, reads ZIPs
with CRC checking, parses relationships/content types, checks exact content-type
lexical insertion, and compares untouched local/central records (central offsets
are allowed to change), order and archive comments. Rust producer gates cover
exact no-op, typed duplicate/missing/stale refusal before output, short reads,
partial sink, output ceiling and exact Part-count limits. The archive oracle
does not independently replay these typed error paths. All 29 mutation probes
reject corrupt archives or report sample/hash/metric/gate claims.

No production API, dependency, unsafe code, native format owner, scheduler or
ambient production I/O changes. The selector registry is 437; the default set
remains 36. The representative coverage index remains 15 categories, 33 mappings,
10 measured and 23 correctness-only mappings; this opt-in package baseline does
not promote semantic owner coverage. Native breadth, repackaging, bounded semantic
append/creation, cold/range sources, codec byte flow and scaling remain open.

The first portable-verifier attempt assumed the earlier ABBA profile receipt
schema and failed on a missing `status_before`. The pinned single-baseline
profile driver records its equality check and final status, not that initial
list. The corrected verifier checks its reported equality, final status against
the last capture, exact driver/helper hashes and source endpoints. Original
receipts and the failed verifier draft remain unmodified.

The sealed exported bundle passes before and after cleanup. Five pre-cleanup
and six post-cleanup portable mutations are rejected, including executable
identity, source custody, argv, extra inventory members and cleanup status.
Cleanup removed only `/tmp/litchi-goal-0444-binaries` (916,671,032 regular-file
bytes); both Cargo target directory identities and GOAL.md's pinned digest
remain unchanged. The final sealed bundle verifies without those binaries.
