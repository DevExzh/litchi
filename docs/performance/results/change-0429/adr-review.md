# 0429 ADR compliance

The accepted ADR tree was read before editing and remains pinned by
`protocol.json` to `c950b6c8be822561b498d7bbe87c460873dcbf49`.

| Constraints | Implementation and evidence |
| --- | --- |
| 0001, 0002, 0010, 0011, 0024: owner and dependency boundaries | ZIP fixed-header refill stays in `soapberry-zip`. Provider adapters, file staging, clocks and journals stay in the standalone performance tool. No production dependency or public API is added. |
| 0003: snapshot/edit/publication ownership | Existing source-backed APIs and cross-copy publication gates remain in use. Native payload handles outlive package owners; checked caller gauges remain observable after those owners drop. |
| 0005: explicit I/O, budgets and measurement scopes | Caller range adapters bound reads and record checked counts/histograms; no networking is introduced. Independent source/destination budgets, cache limits and one-worker execution remain explicit. API clocks exclude setup and observers; RSS remains process-wide. |
| 0006: preservation, typed refusals, malicious input | ZIP short reads assemble complete fixed headers and reject truncated tails. Existing ZIP metadata/signature/ZIP64 bounds remain active. Synthetic exact output gates and native payload/descriptor oracles are required; original shapes refusal fixtures remain untouched. |
| 0008: reproducible verification | Pinned toolchain, retained red and green regressions, source manifests, frozen release protocol, independent report replay and mutation probes bind evidence to concrete bytes. Formal completion is established by receipts and indexes. |

No unsafe code, hidden executor, global cache, ambient production filesystem,
network client, normalization, repair or guessed edit is introduced. This batch
makes no causal performance or global-goal completion claim.
