# 0359q — DOC body toggles defer CHPX publication

- **Status:** implemented on `uo2/doc-toggle`; parent merge pending.
- **Date:** 2026-10-22.
- **Base:** `81323113be7e166d1b37ea185e9fd712d8748a62`.
- **Input:** follow-up 0359a-F1 in `(deleted-coord-artifact)` (the leaf
  body-toggle cost at 1,000 paragraphs).
- **Scope:** `litchi-doc`'s existing body transaction and its W5 probe. No
  public editing family or public type was added.

## Change

`RevisionEditor::set_character_override` used to clone the complete DOC
editor, rebuild and publish the CHPX FKP, and commit the package for every
paragraph toggle. It now clones only the semantic CHPX vector, validates and
transforms that clone, then installs it and marks the editor dirty. The FKP
rebuild and stream publication happen once at the next serialized-stream
dependency or final render. A failed transform is discarded before any editor
state changes, while text and table operations flush the pending CHPX through
the existing commit boundary. The editor remains clone-first and atomic, and
the signed/protected-source and poison/discard paths still run through the
same commit and finish guards.

The probe adds `doc/batch1k_leaf_write`, which stages 1,000 leaf body toggles,
commits once, and writes the complete package with `Package::write_to`. The
`doc/leaf_write` supplies the ordinary one-toggle commit-and-write control, and
the existing facade coalescing test continues to pin one facade commit.
Direct format staging clears a retained render until commit; the final commit
still uses the configured retention ceiling.

## Correctness

The new mixed body-format/table-text test covers a caught invalid selector,
unchanged pending bytes after that refusal, read-after-write table access,
DOC reopen and run formatting, inverse replay, and an exact no-op. The full
`litchi-doc` library suite passed 1,229 tests (2 ignored), and
`litchi-ops`'s `doc_unified_ops` suite passed all 13 tests.

## Measurement

The release probe was built with one Cargo job and run pinned to CPU 29. Both
legs used the same source fixture, three measured rounds, and the same
`LITCHI_PROBE_ONLY=doc/batch1k_leaf_write` and
`LITCHI_PROBE_BATCH_ROUNDS=3` settings. The host was shared with other builds,
so these numbers are diagnostic evidence rather than a native qualification
claim.

| Full leaf commit plus `write_to` | Before | After | After / before |
|---|---:|---:|---:|
| Median (µs) | 1,459,760.103 | 235,103.453 | 0.161 |
| Minimum (µs) | 1,459,383.830 | 235,042.851 | — |

The output SHA-256 was
`9c1cc5368f99a988f1280446fcf43ec9485a1a88ff6e7245b26038676718731a` on both
legs. The before binary was the unmodified DOC implementation with the probe
route (`c945a4b6cdfc6c8c96e142e3d24e02c6bff10a5e83746bbed566cbcef1ea8544`,
13,339,296 bytes); the final after binary was
`4a0f55238da067e8a1c11d3492cd2cf821ff3292f88f147b9906c63a50e48646`
(13,339,328 bytes).

Both binaries came from this source build (with the repository memory-capped
Cargo shim on `PATH`):

```text
cargo build --offline --locked --release -p litchi-ops \
  --features doc,xls,rtf --example legacy_w5_source_probe
```

The reproducible command, after the release build, was:

```text
LITCHI_PROBE_ONLY=doc/batch1k_leaf_write \
LITCHI_PROBE_BATCH_ROUNDS=3 \
taskset -c 29 target/release/examples/legacy_w5_source_probe 3
```

The reduced JSON/log captures and the source diff used for the temporary
before build remain outside the repository:

```text
(deleted-coord-artifact)
  before-leaf-write-3.json       231 B  sha256 8617224f10d004d7b6effea6436e77c8be44c2f90077f701385a8561192de04e
  before-leaf-write-3.log         42 B  sha256 88363bee6c45588b3aab28563872ae463caa803fec68be6750026b7c6e240f59
  after-leaf-write-final.json    223 B  sha256 a821d56037df4d0edd34612358d25b6dc4df8df0cd236c7fd5c575c4e65f75cc
  after-leaf-write-final.log      42 B  sha256 654244963569b73a278ae8e85de84cc3104a7a8585d9696731ce62f194ed4514
  change.patch                 10,887 B  sha256 3fa783b7d2cc9febe4e7bffe38d0cf8122896182099f4d14e951f3b61ace5a6c
```

No large evidence or executable is committed. These shared-host wall-clock
numbers are diagnostic only; native qualification on an isolated host remains
required before this performance repair is accepted as qualified.

## Deterministic profile

The before executable was rebuilt from the base source at `81323113be7e` and
the after executable from this change at `160a00559bb7`. Both included the
same probe routes and were run with one harness sample under Valgrind 3.26.0.
Callgrind instruction counts (`Ir`) and Massif peak heap totals were collected
with the same command shape for each leg. Massif used `--pages-as-heap=no`;
the peak is heap plus allocator overhead, excluding stacks.

| Route | Before Ir | After Ir | Ir ratio | Before peak | After peak | Peak ratio |
|---|---:|---:|---:|---:|---:|---:|
| `doc/batch1k_leaf_write` | 21,183,223,373 | 6,436,404,593 | 0.304 | 113,171,296 B | 14,275,296 B | 0.126 |
| `doc/leaf_write` | 5,426,409,220 | 1,737,517,721 | 0.320 | 104,778,760 B | 14,012,920 B | 0.134 |

The `doc/leaf_write` shared-host wall-clock medians were 3,106.697 µs before
and 3,684.751 µs after. The historical instruction and peak rows above are
whole-process observations that include the unselected batch seed, so they do
not establish ordinary-route CPU or heap changes; the wall-clock difference is
not a qualification result. Both routes produced the same output SHA on both
legs. The route-local ordinary attribution is corrected in
[0359qa](0359qa-doc-body-toggle-profile.md).

The compact profile summary is outside the repository at
`(deleted-coord-artifact)/profile-summary.txt`
(1,064 bytes, SHA-256
`5c66fa441092b06e6d8355c5a542eddc0ee3db3dcffbee148051092e4db9efd1`). The
matched profile binaries are also external: `before-with-single-route-legacy_w5_source_probe`
(`d44a1b17a4377b87f88488def75ec158bbd425b4d849e4f6e969c288ae60a58d`,
13,343,600 bytes) and `after-with-single-route-legacy_w5_source_probe`
(`635d421259b36faf33141be8c8efb63b51f47736c33114b975e0609c05ccd4fc`,
13,343,824 bytes).

The profiling commands were:

```text
LITCHI_PROBE_ONLY=doc/batch1k_leaf_write \
LITCHI_PROBE_BATCH_ROUNDS=1 \
valgrind --tool=callgrind --callgrind-out-file=<leg>-batch.callgrind <binary> 1
LITCHI_PROBE_ONLY=doc/leaf_write \
LITCHI_PROBE_BATCH_ROUNDS=1 \
valgrind --tool=callgrind --callgrind-out-file=<leg>-single.callgrind <binary> 1
LITCHI_PROBE_ONLY=doc/batch1k_leaf_write \
LITCHI_PROBE_BATCH_ROUNDS=1 \
valgrind --tool=massif --pages-as-heap=no --massif-out-file=<leg>-batch.massif <binary> 1
LITCHI_PROBE_ONLY=doc/leaf_write \
LITCHI_PROBE_BATCH_ROUNDS=1 \
valgrind --tool=massif --pages-as-heap=no --massif-out-file=<leg>-single.massif <binary> 1
```

## Validation

```text
cargo fmt --all --check
cargo test --offline --locked -p litchi-doc --all-features --lib
cargo test --offline --locked -p litchi-ops --all-features --test doc_unified_ops
cargo clippy --offline --locked -p litchi-doc -p litchi-ops \
  --all-targets --all-features -- -D warnings
```

All commands passed. The final release probe also passed its byte-for-byte
before/after output assertion.

## Attribution correction

The historical profile rows in §Deterministic profile are retained, but they
are not ordinary-route counts. The `doc/leaf_write` Callgrind and Massif legs
still executed the unselected expected `doc/batch1k_leaf_write` call before
the selector was checked. That 1,000-toggle seed dominates the old whole-
process observation and its 0.320 ratio; the old Massif peak is likewise a
batch-seed peak. The route-local attribution and the selector repair are
recorded in [0359qa](0359qa-doc-body-toggle-profile.md).
