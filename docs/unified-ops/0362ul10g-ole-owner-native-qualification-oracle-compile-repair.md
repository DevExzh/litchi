# 0362ul10g — OLE owner oracle compile repair

- Date: 2026-10-24.
- Parent packet: 0362ul10f.
- Status: consumer-only compile repair; native execution, timing and paired
  qualification remain pending coordinator review.

The correctness oracle had not kept up with the current DOC/XLS APIs.  It now
borrows the archive bytes accepted by `from_bytes`, owns directory metadata
before opening a stream through the mutable CFB reader, keeps the worksheet
value collections alive while walking their worksheet projections, and uses
the same borrowed archive inputs in the retained-history and allocator routes.
The production baseline and candidate source trees were not changed by this
repair.

The first locked diagnostic against the prepared external drivers was refused
by Cargo before compilation: the reviewed `bc5a133f...` lock is stale for the
driver manifest/source graph.  The generated disposable locks differed only
in the harness root's direct `sha2` edge for the baseline and, for the current
candidate, the source's `litchi-odraw` → `litchi-core` edge as well.  The pinned
lock was not modified and no `--locked` failure is treated as a build result.

Using fresh external targets and the generated disposable lock only to isolate
consumer compilation, both endpoint/API combinations reached a clean oracle
check after the repair:

```text
baseline: cargo check --release --offline --no-default-features \
  --features legacy-owner-api,legacy-property-api --example ole_owner_oracle
candidate: cargo check --release --offline --no-default-features \
  --features current-api,w5-counting-allocator --example ole_owner_oracle
```

Both commands used `CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, the repository
memory-capped Cargo dispatch, and distinct targets below
`/home/zhuhe/.cache/litchi-api-contract/ole-owner-oracle-repair-check-20261024/`.
The baseline completed in 50.40 seconds and the candidate in 75 seconds; these
are compile diagnostics only, not timing or performance measurements.  The
logs are external at
`/home/zhuhe/.cache/litchi-api-contract/ole-owner-oracle-repair-check-20261024/`:
`baseline-legacy.log` (5,237 bytes,
`3cb98f0d2dc7b9e618b7336c68392be3603202b6b97ce7bdb77c02974c097d36`) and
`allocator.log` (6,012 bytes,
`b0a938487eb64e5f340a70cb38eff7732d374830922d8e4e038ec0b101b32293`).  The
disposable generated lock identities were baseline
`450f43c6b74536ef552bd5e9a19e321a3c0dea745f645fecc0aa9abe37868e14` and
candidate `5a07adc1d62dea06dbac60986faa020007882506523dde49e4605708e1276aaf`;
the prepared pinned lock remains `bc5a133f0668142d18d637b2b196097e855c8890697cb3bbf43aeb937145f04a`.
The source oracle hash is
`60b506c1bfc277c260933730ce2f754e59f4745d40c753ffa95900b49d276daf`.

Source-only checks performed:

```text
rustfmt --edition 2024 --check tools/perf-baseline/examples/ole_owner_oracle.rs
git diff --check
```

The preparation and manifest oracle pins were updated to the repaired source
hash.  A new source-bound preparation receipt and any locked build attempt
must be generated after this commit; this packet does not claim a native build
or acceptance.
