# 0362ul157b — complete the typed graph test migration

The coordinator's first focused Rust run on integration source
`86a12758d05d5bbefb8ece095ff49b3e990e533d` finds a test-consumer error in
157. `ReferenceGraph::incoming()` returns an optional iterator; comparing
that option directly with `None` requires the iterator to implement
`PartialEq`. The old vector-based forwarding route hid that distinction.

The test now uses `is_none()` to assert exactly the same missing adjacency.
No production API, iterator implementation or graph behavior changes. This
completes the consumer migration without introducing an alias or adding an
irrelevant equality operation to the iterator.

The failed runner packet is
`/home/zhuhe/code/litchi-worktrees/.coord/contract-rust-C86a12758-20261025/`.
Its 1,142-byte log has SHA-256
`4cc506d3fb9ac3e133772a80a2989cd4257b82328fd99b6ef210e46bf6fc94ce`.
Root verifies the failed terminal service, exact log and target-trace
identities, unchanged source and locks, and empty owned process samples.
The IWA library row failed at compilation; the PPT owned-record row did not
run. The target ended at 2,452,716 bytes with no resource-guard failure.

The corrected committed source requires a fresh focused runner receipt.
The failed packet remains unchanged, and no pass is inferred from source
checks or the one-line fix. Full candidate qualification remains open.
