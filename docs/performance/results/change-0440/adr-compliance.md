# Candidate ADR obligations

The accepted ADR tree is unchanged from the prior read:
`c950b6c8be822561b498d7bbe87c460873dcbf49`.

| ADR | Candidate obligation and evidence |
|---|---|
| 0001, 0004 | Private parser representation only; ordinary API and returned owned values remain unchanged. |
| 0002, 0010, 0011, 0023, 0024 | ODP retains its XML grammar ownership; no dependency or physical package API changes. Run crate boundaries. |
| 0003 | No snapshot, edit, publication, patch or conflict behavior changes. Existing append gates retain exact source/no-op sharing, forward/inverse patch and stale-source checks. |
| 0005 | Remove repeated owned namespace URI copies while retaining a per-element cache. The resolver borrow ends with the cache; it cannot outlive or race mutable reader state. No hidden cache, executor, ambient provider or unsafe code. Require fresh before/after latency, allocator and RSS evidence. |
| 0006, 0008 | Preserve exact namespace matching, lazy error order and typed failure messages. Run independent direct-lookup and harvest oracles, the full ODP suite, harness suite, strict owner Clippy, rustdoc and formatting. Retain existing dependency lint failures separately. |

No new architecture or ADR exception is proposed. End-to-end package opening,
validation, readback, recompression, publication and retention boundaries are
unchanged. This experiment does not establish bounded existing-document append
memory, raw passthrough, cancellation coverage, native rendering, or scaling.
