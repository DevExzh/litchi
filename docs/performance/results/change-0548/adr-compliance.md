# 0548 ADR compliance

The accepted ADR/index inventory remains byte-identical to the previously read
inventory, as bound by adr-manifest.json. The experiment changes one private
CFB collector and adds a public-API benchmark example. Retention depends on the
frozen semantic, resource and performance gates.

| Constraint | Evidence / boundary |
| --- | --- |
| 0001 priorities and API layers | No public API, dependency or unsafe-code change. Exact errors and bounded malformed work precede latency. |
| 0002, 0010, 0011, 0024 ownership | Validation stays in the low-level CFB owner; no duplicate validation moves into XLS, OOXML or the facade. |
| 0003 immutable snapshots and atomic publication | Collection completes before physical sector claims. No edit, commit, conflict or snapshot behavior changes. |
| 0005 performance and budgets | Fresh paired native, allocation and instruction evidence; unchanged initial reservations and visited zero-fill. Checked checkpoint counters bound speculative cycle work; guard timings test malformed inputs independently. |
| 0006 preservation and safety | Authoritative replay preserves ordered corruption diagnostics. FAT/MiniFAT scratch, ownership checks and reconciliation remain in place. Existing diagnostic string allocation is explicitly outside the no-scratch-growth claim. |
| 0008 verification | Source/binary-bound receipts, exhaustive Rust differential tests, public malformed-input oracles and final OLE quality gates. No unavailable sanitizer or native-Office campaign is claimed. |
| 0026 OLE directory binding | CFB retains physical-chain validation; concrete formats and shared OLE metadata ownership are unchanged. |

No ADR exception or replacement is required. The 0547 finite model supports
the design but does not replace implementation tests or measured admission.
The final decision and quality summary determine which source is retained.
OLE2/OOXML optimization remains active; ODF is deferred and iWork excluded.
