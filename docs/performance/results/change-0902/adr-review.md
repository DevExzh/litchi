# 0902 candidate ADR review

This is a local implementation specialization within `xml-minifier`; it changes
no public API, dependency edge, execution context, or architectural boundary.
The accepted ADRs remain unchanged from the already-reviewed production baseline.

| Constraint | Candidate treatment | Evidence |
| --- | --- | --- |
| 0001 priorities; 0006 validation and preservation | Both complete source-policy XML audits remain mandatory, original before replacement. No repair, normalization, policy weakening or limits change. | Source review; original/replacement exact-error and byte-limit parity tests; full XML and Office tests |
| 0002 / 0024 ownership and dependencies | Existing low-level XML owner; no peer-format dependency or exported archive types. | Two-file patch and crate-boundary check |
| 0003 immutable edits and atomic publication | No snapshot, edit, commit, conflict, patch or publication mutation. | Existing OPC/XLSX/DOCX/PPTX suites and workload preservation oracles |
| 0005 I/O, memory and evidence | No I/O, allocator, caching, executor, thread or ambient provider additions. Warm paired native and separate allocation captures; no cold/concurrent/general claim. | Frozen method, command receipts and exact-source manifests |
| 0008 verification state | Candidate adoption depends on the frozen end-to-end gates; rejected code is restored with evidence retained. | Decision after capture |
| 0010 / 0011 physical package ownership | No archive entry or package rewrite logic change. Unchanged member bytes and source/sink identities remain checked. | Workload semantic/resource validation |
| 0032 snapshot memos | No audit proof memo or cached acceptance is introduced. Every invocation still performs both required audits when upfront search is disabled. | Source review; strict half-boundary, malformed XML, namespace, Unicode and UTF-8 tests |

No parallelism, unsafe code, SIMD, resource amplification, or new cancellation
behavior is introduced. The existing window proof/replay invariants from changes
0747 and 0750 remain intact. The `minify_xml` fuzz target covers unchanged slice
and stream auditors, not the changed replacement entry point; this batch relies
on focused differential replacement tests and existing Office corpus tests for
that branch, and makes no new fuzz-coverage claim.
