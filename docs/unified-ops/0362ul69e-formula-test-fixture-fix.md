# 0362ul69e — Formula raw-policy test-fixture compile fix

Date: 2026-10-25. Parent: 0362ul69d. Status: source repair complete; the
coordinator still owns the next guarded validation attempt.

The first post-0362ul69d guarded Formula run compiled the library and then
stopped in the test harness. The new ruled-depth and append-retention tests
passed owned AST nodes into temporary slices and reused those nodes, producing
two `E0382` moved-value errors for `double` and `bad_font`. The failure was
limited to test setup; no production code ran and no resource or provenance
guard failed.

The external f5 receipt is:

    (deleted-coord-artifact)/01-formula-all-targets-tests-attempt-formula-all-targets-post-fix.log
    1860 bytes, SHA-256 6ec3b437cc4c7ef542c1e1a1bb93fa5ceef311e68454818733c911450acd0cc8

The repair clones each fixture only for its first temporary slice, leaving the
second call unchanged. This preserves the exact test inputs and does not alter
the writer, depth policy, append rollback, output, or semantic behavior.

The guarded rerun remains required for the complete Formula all-target suite,
the focused depth/output/append-prefix cases, affected DOC/DOCX/ops/facade
consumers, and strict Clippy. Exact 4100 and original tool locks remain
hydrated and byte checked; no lock generation, native, rustdoc, broad ops, or
performance command is part of this return.
