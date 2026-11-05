# 0359qb — corrected ordinary DOC body-toggle qualification

- **Status:** matched heap and Callgrind evidence recorded; native blocks
  retained as diagnostic rejects because the host was busy.
- **Date:** 2026-10-22.
- **Scope:** `doc/leaf_write`, the one-toggle commit-and-write route. This
  record uses the selector repair from 0359qa, so the unselected 1,000-toggle
  batch seed is absent from both legs.

## Source control

The before executable was assembled from the pre-change implementation at
`81323113be7e166d1b37ea185e9fd712d8748a62`, with only the corrected probe
copied from `428bf4d44a237bc56207b07a220a3d1da3043f04`. The after executable
used the implementation from `160a00559bb7edb905b1315753e33aedc74c3733`
and the same probe source. This is an explicit two-source assembly for a
matched harness; the before leg is not described as a pristine 813 checkout.

The implementation file hashes were:

| Assembly | `body_text.rs` | `tracked_revision/package.rs` |
|---|---|---|
| Before | `ae044d85d99a987ac34d08c8f6d8f605e85f5443bb66573964a88a422e047d46` | `613beaf162033c6153906ef2b1299d21e0ece58b4283a43400924c2503dbe40a` |
| After | `b6b23536a8a73f44f49ae1d961784b92e8fa68524f0c744875b4934f78afccf3` | `722d5ce021ce14bc1228ead31e6992d495e6054974edbc15bacab330c05d3eb4` |

Both used probe hash
`9e1cd649268ad5eefe4314b76493f61cd047bf6ec4114c7e7872b2182b1af689`.
The release build command was the same for both legs, with one Cargo job:

```text
cargo build --offline --locked --release -p litchi-ops \
  --features doc,xls,rtf --example legacy_w5_source_probe
```

The protected external binaries were 13,346,048 bytes before
(`0c4392e6cbf5a65b5e3901fdc860cec80ce29518b11e47b70620641df405c92a`) and
13,346,080 bytes after
(`cfdf1e3040b9b742551d80cd583fc28de4f366f9f9bc9fb7198e91aa3352384f`). A
single stable runner path and argv0 was swapped between legs:

```text
/home/zhuhe/code/litchi-worktrees/.coord/doc-toggle/tmp/runner/legacy_w5_source_probe
```

Every selected route call asserted its output against the expected bytes. Both
legs produced output SHA-256
`6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11`.

## Matched heap

The Massif command used `--pages-as-heap=no`, one harness round, and the same
`LITCHI_PROBE_ONLY=doc/leaf_write` and
`LITCHI_PROBE_BATCH_ROUNDS=1` environment on both legs. The peak tree was
under `routes::doc_leaf_write` in each capture and contained no
`doc_batch1k_leaf_write` frame.

| | Before | After | After / before |
|---|---:|---:|---:|
| Heap | 8,182,107 B | 7,573,945 B | |
| Allocator overhead | 46,541 B | 46,207 B | |
| **Peak total** | **8,228,648 B** | **7,620,152 B** | **0.926052** |

The external captures are `qualified-ordinary-before.massif` (110,488 B,
SHA-256 `1461820dd6ac7b7c99e3b4232925068ab11e8765ea2cab10a23e12dd76668b75`)
and `qualified-ordinary-after.massif` (199,478 B,
SHA-256 `b06f68779f5cdf3923254a47fa95b86bd8efec4f13211cff249b46a76f755b76`).

## Corrected route-local instructions

Callgrind used one harness round, so the selected route node contains four
calls: expected output, two warmups, and one measured call. The corrected
route-local inclusive `Ir` counts were:

| Operation | Calls | Before Ir | After Ir | After / before |
|---|---:|---:|---:|---:|
| `routes::doc_leaf_write` | 4 | 256,566,637 | 258,771,387 | **1.008593** |
| `Package::edit` | 4 | 87,950,819 | 92,226,703 | 1.048617 |
| package open (`from_bytes_with_options` / `from_owned_bytes`) | 4 | 60,176,893 | 59,526,823 | 0.989197 |
| `set_character_property_override` | 4 | 32,187,392 | 22,491,847 | 0.698778 |
| `Edit::commit` | 4 | 73,309,585 | 80,778,169 | 1.101877 |
| `Package::write_to` | 4 | 2,284,447 | 2,413,559 | 1.056518 |

The Callgrind captures are `qualified-ordinary-before.callgrind` (313,941 B,
SHA-256 `63e21a13fe16eb81755ad20b1c1bcda30f1d1e9d0c7b0b752731fcaee1042ee3`)
and `qualified-ordinary-after.callgrind` (313,464 B,
SHA-256 `b9c3e9303ec771de4c05e278a25c39bcd6419ac5ae71c59c479bdacf4d7c6165`).

The earlier 0359qa pair still reports the valid 0.988216 route-local value
for its own binaries. Its before/after probe binaries predate the selector
repair; this corrected-probe pair is reported separately because changing the
probe changes the `litchi-ops` executable and its code generation. The
component increases above are measured observations against the original
implementation; they are a follow-up optimization target and do not waive
the large-batch improvement or any atomicity test.

## Native attempts

Two short ABBA blocks used the same runner, route selector, three measured
rounds, output assertion, and environment. They are retained as rejects: the
host load stayed between 117.52 and 126.85, with unrelated Cargo, rustc and
ninja work active during the runs. CPU29 block N1 had medians 3306.838,
3470.709, 3400.668, 3245.517 µs (A average 3276.178, B average 3435.689,
B/A 1.048688). CPU19 block N2 had medians 3314.258, 3428.518, 3498.488,
3417.757 µs (A average 3366.007, B average 3463.503, B/A 1.028965). Every
receipt has the same output SHA above. These are diagnostic observations, not
an isolated native qualification.

The eight JSON receipts and their matching logs remain outside the repository
under `/home/zhuhe/code/litchi-worktrees/.coord/doc-toggle/` as
`native-abba-a1-*` and `native-abba-a2-*`. The machine-readable summary is
`docs/unified-ops/results/0359qb/qualification.json`.

## Reproduction boundary

Massif and Callgrind were run under a 20 GB systemd scope with the repository's
memory-capped Cargo shim on `PATH`; no binaries or large captures are committed.
Quiet native qualification remains required. The next repair should reduce the
ordinary-route `Edit::commit`/write overhead while preserving the deferred
CHPX publication, mixed-operation flushes, exact output, and batch result.
