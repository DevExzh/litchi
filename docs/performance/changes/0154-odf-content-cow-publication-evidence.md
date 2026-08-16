# Change 0154: ODF content-COW publication evidence

Date: 2026-08-16

Status: accepted for the prepared in-memory ODT/ODS/ODP publication boundary.
No end-to-end, allocation/RSS, physical-I/O, decompression, cold-cache,
filesystem, real-producer, or iWork claim is accepted.

## Scope

Commit `6218ecab331e2b82e2a9fd762d5ead4d6098a451` adds six opt-in
selectors to the standalone harness:

- `odt_content_cow_owned_rebuild`
- `odt_content_cow_positional`
- `ods_content_cow_owned_rebuild`
- `ods_content_cow_positional`
- `odp_content_cow_owned_rebuild`
- `odp_content_cow_positional`

The current matrix therefore has 301 selectable names. The historical default
remains 36 cases / 198 records.

Each family uses its real semantic edit path outside timing: one middle ODT
paragraph replacement, one middle ODS scalar-cell replacement, or one ODP
text-box addition. The candidate `content.xml`, family owner, positional
archive index, output length, and expected digest are prepared before the
timer. The control calls the family-neutral owned rebuild. The candidate calls
the bounded source-positional content publisher. Both then emit to the same
fixed 16 KiB non-seek hashing sink, which retains no complete output.

This is an intentionally asymmetric product-path comparison. The owned rebuild
materializes and retains one complete output during the timed call; the
positional publisher emits directly and reports zero retained complete
candidate bytes. The sink hashes both outputs symmetrically. Candidate
generation, package open/indexing, semantic reopen, logical source replay, and
all refusal gates are outside `elapsed_ns`.

## Correctness and evidence gates

Every retained record verifies:

- exact candidate `content.xml` and family semantic reopen, including media;
- unchanged package member inventory;
- for the positional output, raw identity of every untouched local member and
  central record apart from the necessarily relocatable local-header offset,
  plus unchanged physical local-header order and central-directory order;
- exact no-op byte identity;
- typed one-byte-under replacement and output limits;
- pre-output cancellation and source immutability;
- the fixed 16 KiB sink window, zero retained output, exact output length and
  digest; and
- one logical `ReadAt` replay per retained positional sample, classified into
  `content.xml`, other untouched members, and `Pictures/*` compressed ranges.

The source counters are logical byte-range overlap over an immutable in-memory
`ReadAt`. They are not physical reads, syscalls, decompression work, device or
filesystem I/O, or network requests. The complete B1 and B2 source-evidence
objects are byte-for-byte identical.

Final proportional verification was:

- focused ODF selector/parser and all-six execution tests: 2/2;
- strict all-target Clippy with warnings and deprecations denied;
- rustfmt, diff check, and the 64-package/238-edge boundary checker;
- complete harness suite: 122/122 before the final review hardening; and
- final focused 2/2 plus strict Clippy after the review-only schema naming,
  member-order proof, and overlap-assertion corrections.

Two independent read-only reviews returned SAFE on the final revision.

## Clean release ABBA

The final release binary has SHA-256
`4a45436ce46331ff96b5ab8549cdbfd5efe25f0540914a9d7f28c964a17c04de`.
It was built inside a clean detached worktree at `6218ecab3` and run on CPU 2.
All reports record `git_worktree_dirty=false`, one affinity-visible CPU, Rust
1.95.0, Linux 6.8.0-101-generic, and the system allocator. The strict order
was `A1 owned, B1 positional, B2 positional, A2 owned`; every record used 20
warmups and 100 retained samples, for `4 * 3 * 100 = 1,200` samples.

The individual raw rows deliberately retain
`publication-only correctness evidence; no speedup claim`; acceptance exists
only after this four-leg aggregate review. Owned and positional output hashes
are individually deterministic. ODS happens to produce identical bytes;
ODT/ODP use different valid ZIP framing, so the matched comparator is exact
candidate content, member inventory, semantic/media reopen, and deterministic
per-path output rather than cross-implementation byte identity.

P50 publication times and paired improvements are:

| Family | A1 owned | B1 positional | B2 positional | A2 owned | A1 -> B1 | A2 -> B2 |
|---|---:|---:|---:|---:|---:|---:|
| ODT | 224.782 ms | 7.669 ms | 7.559 ms | 224.546 ms | 96.588% | 96.634% |
| ODS | 223.905 ms | 8.179 ms | 8.172 ms | 224.277 ms | 96.347% | 96.356% |
| ODP | 224.015 ms | 7.695 ms | 7.589 ms | 222.658 ms | 96.565% | 96.592% |

The paired distribution results agree beyond p50:

| Family | P95 A1 -> B1 / A2 -> B2 | P99 A1 -> B1 / A2 -> B2 | Mean A1 -> B1 / A2 -> B2 |
|---|---:|---:|---:|
| ODT | 96.544% / 96.721% | 96.516% / 96.754% | 96.581% / 96.642% |
| ODS | 96.279% / 96.513% | 96.306% / 96.534% | 96.340% / 96.378% |
| ODP | 96.583% / 96.612% | 96.615% / 96.679% | 96.558% / 96.593% |

Same-implementation p50 drift was -0.105%, +0.166%, and -0.605% for
ODT/ODS/ODP controls, and -1.441%, -0.082%, and -1.379% for the candidates.
The largest absolute p50 drift was 1.441%. Raw standard deviations and
two-sided Student's t mean-confidence intervals are retained in the summary
and raw reports; both pair directions agree by far more than the observed
same-implementation drift.

## Logical source evidence

Each positional sample replays the exact prepared publication after timing:

| Family | Calls | Bytes | Ordinary calls / bytes | `content.xml` bytes | Other untouched bytes | `Pictures/*` bytes |
|---|---:|---:|---:|---:|---:|---:|
| ODT | 611 | 16,789,267 | 525 / 16,783,466 | 950 | 1,090 | 16,782,376 |
| ODS | 597 | 16,793,240 | 523 / 16,782,726 | 6,383 | 350 | 16,782,376 |
| ODP | 611 | 16,789,117 | 525 / 16,783,443 | 823 | 1,067 | 16,782,376 |

The values are identical in all 100 B1 samples and all 100 B2 samples for each
family. They show which compressed source ranges the publisher requests; they
do not establish physical-I/O or decompression savings.

## Artifacts

The [machine-readable summary](../results/odf-content-cow-abba-0154-summary.json)
has SHA-256
`210e9fa7789075bd0516dbfb2b3b2fb4250e346d121294a77b96a59379f21e45`.
The complete raw reports are retained as:

- [A1 owned](../results/odf-content-cow-a1-owned-0154.json.zst), compressed
  SHA-256 `aa228940a78564063bc3ec35405d66b12d616920088f56044aab21747281d80d`;
- [B1 positional](../results/odf-content-cow-b1-positional-0154.json.zst),
  compressed SHA-256
  `4c968478ffc3fce7e6c5f7ec7d8bcf47bbcf6aa098acf4a98c0fe293a528ccb5`;
- [B2 positional](../results/odf-content-cow-b2-positional-0154.json.zst),
  compressed SHA-256
  `633770e428d992d68f2582af5048ac68536b07323962a35efb8968f048c2f96c`;
- [A2 owned](../results/odf-content-cow-a2-owned-0154.json.zst), compressed
  SHA-256 `493ecd91bd01625c11e41b6745259ec5d9a941db1218bca9cac6dcd442395851`.

The summary also retains raw sizes/hashes, corpus and output hashes, exact
statistics, confidence intervals, counters, protocol, and claim boundaries.

## Accepted boundary and remaining gaps

The accepted result is that, for these deterministic approximately 16 MiB
media-rich corpora and already prepared owners/candidates, the source-
positional `content.xml` publisher completes the named in-memory publication
call about 96.3%-96.8% faster across p50/p95/p99/mean in both pair directions
while retaining the stated correctness and logical-source gates.

This does not include semantic edit construction, archive open/indexing,
reopen, validation gates, or filesystem output. No allocation count, allocated
bytes, peak heap, RSS, cold-cache, physical I/O, syscall, decompression,
recompression, energy, throughput, concurrency/scaling, real-producer, broad
ODF CRUD, or iWork conclusion follows. Those remain separate evidence gaps.

## Reproduction

Build from a clean worktree, pin the resulting release binary to one CPU, and
run the four legs in order:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml

taskset -c 2 tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 20 --samples 100 \
  --case odt_content_cow_owned_rebuild,ods_content_cow_owned_rebuild,odp_content_cow_owned_rebuild \
  --json target/perf/odf-content-cow-a1-owned.json

taskset -c 2 tools/perf-baseline/target/release/litchi-perf-baseline \
  --warmup 20 --samples 100 \
  --case odt_content_cow_positional,ods_content_cow_positional,odp_content_cow_positional \
  --json target/perf/odf-content-cow-b1-positional.json

# Repeat the positional command for B2, then the owned command for A2.
```
