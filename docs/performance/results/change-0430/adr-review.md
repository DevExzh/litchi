# ADR compliance

The accepted ADR tree is unchanged at
`c950b6c8be822561b498d7bbe87c460873dcbf49`. The coordinator previously read
every accepted ADR and reread the index while checking this batch. No Rust,
manifest, lockfile, API, or dependency change is made here.

| Constraint | Evidence and next implementation boundary |
|---|---|
| 0001 / 0006: correctness, preservation, typed refusal first | Current correctness gates remain active in both profile reports. Profiling does not authorize weakening candidate, CRC, XML, source, or dependency validation. |
| 0002 / 0010 / 0011 / 0024: grammar and package ownership | ZIP must own any compressed entry representation and framing; OPC must own an opaque part-transfer operation. PPTX must not acquire ZIP dependencies or raw entry IDs. |
| 0003: immutable snapshots and checked publication | Any future transfer must retain source identity/version authority and recheck it at publication. A matching logical payload or URI alone is insufficient authority. |
| 0005: explicit providers and bounded resources | Captures use explicit bytes/file providers and CPU affinity. Files are warm. Future transfer must reserve compressed staging or use bounded source reads, with output limits and cancellation on failure paths. |
| 0005: representative measured performance | Two profiles identify a synthetic media-rich publication bottleneck. No latency, allocation, RSS, native, parallel, or global speedup result is claimed. Normal matched before/after captures remain required for implementation. |
| 0008: honest verification state | Portable evidence replay is separate from Rust correctness gates. The binary is unchanged; prior 0429 Rust results are historical evidence, not newly rerun checks. Existing lint debts remain open. |

No ADR exception or proposed amendment is needed for this evidence work. The
transfer design remains a proposal subject to the existing constraints.
