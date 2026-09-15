# Evidence packet — change 0628

Relationship iteration order reaches no published `.rels` byte, no catalog
verdict, no signature digest and no read order; it did decide which of two
duplicate relationships `Relationships::get_or_add` / `get_or_add_ext_rel`
reuses, which four packages in this repository's corpus trigger. Record:
[`docs/performance/0628-opc-relationship-iteration-order.md`](../../0628-opc-relationship-iteration-order.md).
`performance_claim: none`.

## Provenance

| field | value |
| --- | --- |
| base commit | `2d6fbeaed` (`feat/office-format-completeness`) |
| branch | `perf/0628-opc-relationship-iteration-order` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0628` (removed after commit) |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-2d6fbeaed` (shared, read-only; used only to reproduce the pre-existing test failure) |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd 2026-03-21), valgrind 3.26.0 |
| build | `cargo build --release` for every probe, both legs, same worktree |
| CPU affinity | every measured process pinned to CPU 22 with `taskset` |
| host state | eight agents building and testing concurrently throughout |

The two legs are the same worktree with `crates/litchi-opc/src/rel.rs` reverted
(before) or applied (after); nothing else differed. Probe binary SHA-256:

| binary | before | after |
| --- | --- | --- |
| `corpus_roundtrip` | `ce8306de507a9496…6c2d7bac` | `c0cd314178c3c1e4…04069530` |
| `reuse_selection` | `5414b0c0bcb8f98b…0f4f0ee9` | `219a97912e357e88…95d1b018` |
| `open_save_counts` | `79b167fe1df21879…b5441989` | `338ea36f72eeb7e6…601b5452` |
| `open_save_iters` | `1f21ce0521086330…db0dec8d` | `7880e66adc7f19dd…ad9b3d16` |

## Contents

| path | what it is |
| --- | --- |
| `gates.txt` | the four gates, each with its tail and exit status, plus the reproduction of the pre-existing `source_backed_batch` cancellation failure on the untouched base checkout, plus a summary of the differential checks |
| `differential/corpus-duplicate-relationships.txt` | the static scan: 336 OOXML fixtures, 2,218 `.rels` parts, 6,393 `Relationship` elements, and the four packages that carry a duplicate (Type, Target, TargetMode) group, with the owning member and the competing rIds |
| `differential/determinism-before.txt`, `differential/determinism-after.txt` | `cargo test --test relationship_selection_determinism`, one leg each. Before: 3 passed, 4 failed. After: 7 passed. |
| `differential/roundtrip-before.txt`, `differential/roundtrip-after.txt` | `corpus_roundtrip` over all 336 OOXML fixtures: part count, package and part relationship counts, published byte length and a SHA-256 of the published bytes per package, plus every open and save refusal. `diff` of the two is **empty**. |
| `differential/reuse-before.txt`, `differential/reuse-after.txt` | `reuse_selection` over all 336 fixtures, 16 opens each, 5,258 (owner, type, target, mode) triples. Before: 4 triples returned more than one identifier, listed by name. After: 0. |
| `counts/counts-before.txt`, `counts/counts-after.txt` | `open_save_counts` on four fixtures: positional source requests and bytes for a source-backed open, allocations and allocated bytes for the source-backed open, the eager open and the save, and the published byte count. `diff` of the two is **empty**. |
| `callgrind/{before,after}-{1,11}.txt` | valgrind summaries for the N = 1 and N = 11 isolation pairs of `open_save_iters` on `ConditionalFormattingSamples.xlsx` |
| `callgrind/full-{before,after}-11.txt` | `callgrind_annotate --threshold=99.5` self-cost tables for the N = 11 runs; the record's per-symbol attribution is their difference |
| `probe/` | the complete source of all four probe binaries, with `Cargo.toml`, and `scan_dup_rels.py`, the static corpus scan. Path dependencies point at the 0628 worktree; retarget them to reproduce. |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, for the coordinator to merge |

## Headline numbers

| measurement | before | after | tier |
| --- | --- | --- | --- |
| corpus reuse triples with more than one outcome (5,258 probed, 16 opens each) | 4 | 0 | measured |
| published packages byte-identical across legs (of 327 published, 336 attempted) | — | 327 / 327 | measured |
| deterministic counts on 4 fixtures (requests, bytes, allocations, output bytes) | — | identical, `diff` empty | measured |
| instructions per open + save, `ConditionalFormattingSamples.xlsx` | 194,662,395.8 | 194,669,991.4 | measured |
| that delta | — | +7,595.6 (+0.0039%) | measured |
| paired latency | not run — see the record's *Measured* section for why | | — |

The +0.0039% is code layout and allocator arena drift from recompiling the
crate, not new work: neither `get_or_add`, nor `get_or_add_ext_rel`, nor
`reuse_candidate` appears in either callgrind profile, because neither open nor
save calls them. This host's A/A timing floor is about p50 4%, p99 14%, three
orders of magnitude above the delta.

## Reproducing

```sh
# static scan of the corpus
python3 probe/scan_dup_rels.py /path/to/litchi/test-data

# the probes (retarget probe/Cargo.toml's path dependencies first)
cd probe && CARGO_TARGET_DIR=/some/disk/path cargo build --release
cd /path/to/litchi
taskset -c 22 .../corpus_roundtrip   test-data
taskset -c 22 .../reuse_selection    test-data 16
taskset -c 22 .../open_save_counts   test-data/ooxml/xlsx/sharedhyperlink.xlsx ...
for N in 1 11; do
  taskset -c 22 valgrind --tool=callgrind --callgrind-out-file=cg-$N.out \
    .../open_save_iters test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx $N
done

# the determinism test
taskset -c 22 cargo test -p litchi-opc --test relationship_selection_determinism
```

For the before leg, revert `crates/litchi-opc/src/rel.rs` and rebuild the
probes; nothing else changes.
