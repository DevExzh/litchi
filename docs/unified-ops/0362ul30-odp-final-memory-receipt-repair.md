# 0362ul30 final follow-up — ODP memory receipt repair

Date: 2026-10-24. This source-only follow-up addresses the final review after
`2a20b44488`. The earlier packets and records are preserved; this record binds
a new frozen memory protocol and does not claim Cargo, native, timing,
heaptrack, qualification, or acceptance work.

The review is
`/home/zhuhe/.cache/litchi-api-contract/odp-build-handoff-receipt-review-2a20b44488.json`,
6,638 bytes, SHA-256
`58fd4e507de13ff9a6f5c6f0abf2b2a6a722a87dffb48d0e07d2398ac9565534`.

## Frozen final packet

The preserved build packet remains
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-build-prep-472b670547-repair-11858e697f`.
The final memory protocol is frozen at
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-memory-qualification-2a20b44488-final/memory-protocol.json`:
28,914 bytes, SHA-256
`868313fa19b0fd3652bf84644899406721d8f811641cf327f417eb4f9895b3ec`.
The final handoff binding is
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-build-prep-472b670547-final-2a20b44488/final-handoff.json`:
3,664 bytes, SHA-256
`44617c1befa3fa8ade24b17046221e50dca2697af91c754d3d7fac4d385d999b`.

The prior receipt record listed the protocol as 22,912 bytes with hash
`c638ad86e6a7edcf1532408156684d99060020a75d04a31aa67693a1d5c6658b`. That
record is not edited. The frozen final file's actual identity above is the
authoritative correction.

## Final checks

- The pre-run rechecks generated-build-receipt binary hashes, source revisions,
  source trees, all three lock files (including the ignored baseline lock),
  probe hashes, fixture hash, CPU and exact route arguments, and global-quiet /
  freshness guards.
- `status: measured` is possible only when both legs contain
  `peak_heap_bytes`, `peak_rss_bytes`, and histogram-derived `allocated_bytes`,
  with a successful histogram artifact. Partial metrics produce `incomplete`.
- The source-only negative test log is
  `/home/zhuhe/code/litchi-worktrees/.coord/odp-final-memory-qualification-2a20b44488-final/negative-partial-metrics-tests.log`,
  47 bytes, SHA-256
  `ca7e7743983a6fc3d6b5f59bff03b8e42ec358e675623c0f36c7d01f5447563d`;
  all five rejection cases passed.

The historical +0.5 MiB ODP peak-heap residual remains unresolved pending a
future authorized memory run and review.
