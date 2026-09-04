# Change 0395: OPC source-backed case-fold lookup index

Date: 2026-09-04

Status: scoped production optimization for unmanaged source-backed packages.
The default matrix and public APIs are unchanged.

## Implementation boundary

`SourceBackedPackage::part_index` keeps its exact `PackURI` hash lookup as the
common path. The former bounded `eq_ignore_ascii_case` scan is retained for
small catalogs and all managed opens. Unmanaged normal and validation opens
with at least 2,048 admitted ordinary Parts retain an immutable order vector
of Part positions sorted by the allocation-free ASCII case-fold comparator;
case-insensitive misses then use binary search. No folded-name `String`s are
stored, and the vector costs one `usize` per admitted ordinary Part. The
2,048 boundary is a conservative measured tuning threshold, not a semantic
part-count limit.

The index is fallibly reserved and constructed as part of catalog validation
for validation opens. Managed `ExecutionContext` opens deliberately leave it
absent: they retain the bounded linear fallback so cancellation-aware work
does not gain unreserved retained memory or a non-interruptible sort. Source
iteration order, canonical spelling, freshness-before-lookup, mutable
`OpcPackage`, and the public API are unchanged.

## Corpus, protocol, and provenance

The 0394 harness supplies deterministic stored OPC archives with 256, 2,048,
and 16,384 ordinary Parts, each with a 32-byte payload. The fixed lookup
vector contains exact first/middle/last queries, case-only aliases for those
positions, and first/middle/last misses, repeated 16 times (144 lookups).
Source-backed lookup counters are collected by an independent untimed replay;
the timed loop includes no source setup, query construction, output hashing,
or correctness checks. The final release package used `A1/B1/B2/A2`, CPU
affinity 2, one worker, five warmups, and 30 retained samples per case for
both normal and allocator-enabled binaries. Every latency value below is from
the normal, non-allocator binary opened through unmanaged
`SourceBackedPackage::from_read_at`; allocator-enabled runs supplied
allocation vectors, and their latency is observational only. Validation
constructor runs are correctness-only and do not authorize timing claims.

All final builds and tests used explicit stable Rust/Cargo 1.98.1 because the
repository-pinned 1.95 installation had no Cargo. The final package binds:

* base revision `57c8ed4bd8c02938eb7ef21e0d713c05be062125`;
* final patch SHA-256
  `9f6fef3f8a64a3be76b7dd5653831b8dfdd50e4ad959c477047b124ed6a27f31`;
* final `source_backed.rs` SHA-256
  `af43eafe9f287a7581a5e08ef4240840f967dce4d7764c12a3cd3fe32e493417`;
* baseline source blob SHA-256
  `4abd75652327513eecfb624abc9e7c227c77c298a01b74ebf2bd427ecd6b6cc1`;
* shared corpus-catalog SHA-256
  `6d69a092ee45159ebdb446e9e675ba4bcc652b2b36492cd74e2984b6b3ba00f3`;
* baseline normal/allocator binary SHA-256
  `cddfea6a11680f0b721ca7652b3b13119cacfa8a976e30a6aa2d201197167f41` /
  `883369e7dc6effe6fa31d50aed671a476534755b38e344cc30db2e1d9c1f7e52`;
* final normal/allocator binary SHA-256
  `171aeeafd689800256922c09d5583bc0c2493dd0553ee1ef8245133f2c7067f6` /
  `d7a7514e8b25c0ae591af44acb4019ad4197ac4852b84aa92f469c1ada646f6a`.

The [0395 evidence bundle](../results/change-0395/) contains the compressed
probe and final reports, catalogs, metric summaries, and adjudication.

## Probe and final results

The first unthresholded candidate indexed every unmanaged catalog. Its matched
probe rejected the all-size design: source-lookup p50 changed by `+31.31%`
and `+33.45%` at 256 Parts, versus `-74.50%`/`-74.81%` at 2,048 and
`-96.50%`/`-96.62%` at 16,384 in the two `A1/B1` and `A2/B2` directions.
These are normal, non-allocator `from_read_at` latency observations. Source-
open overhead in that probe reached at most `+4.04%` p50 and `+3.74%` mean
in the same normal binary, while the retained index added exactly one
allocation and `8 * parts` allocated bytes. The 2,048 threshold is therefore
the lowest measured corpus with a beneficial lookup result, rather than an
unmeasured 1,024 guess.

For the accepted thresholded candidate, the normal, non-allocator
`from_read_at` source-backed lookup p50 values and paired deltas were:

| ordinary Parts | normal non-allocator A1 → B1 p50 (ns) | normal non-allocator A2 → B2 p50 (ns) | p50 delta |
| ---: | ---: | ---: | ---: |
| 256 | 21,230 → 21,210 | 21,300 → 21,075 | -0.09% / -1.06% |
| 2,048 | 144,801 → 37,320 | 144,535 → 37,050 | -74.23% / -74.37% |
| 16,384 | 1,414,561 → 48,130 | 1,351,806 → 47,965 | -96.60% / -96.45% |

The corresponding normal, non-allocator `from_read_at` source-open p50
values were:

| ordinary Parts | normal non-allocator A1 → B1 p50 (ns) | normal non-allocator A2 → B2 p50 (ns) | p50 delta |
| ---: | ---: | ---: | ---: |
| 256 | 303,881 → 306,536 | 304,066 → 317,746 | +0.87% / +4.50% |
| 2,048 | 2,443,180 → 2,526,600 | 2,432,011 → 2,525,091 | +3.41% / +3.83% |
| 16,384 | 24,232,600 → 24,571,511 | 23,909,837 → 24,269,030 | +1.40% / +1.50% |

For completeness, final normal, non-allocator `from_read_at` paired latency
deltas are reported in `p50/p95/p99/mean` order. Source lookup (`A1/B1`, then
`A2/B2`) was:

| ordinary Parts | normal non-allocator A1 → B1 | normal non-allocator A2 → B2 |
| ---: | ---: | ---: |
| 256 | -0.09% / -13.06% / -3.52% / -0.21% | -1.06% / -18.37% / -21.06% / -2.95% |
| 2,048 | -74.23% / -72.10% / -71.42% / -74.16% | -74.37% / -75.28% / -74.07% / -74.49% |
| 16,384 | -96.60% / -96.60% / -96.60% / -96.60% | -96.45% / -96.11% / -96.26% / -96.43% |

Normal, non-allocator `from_read_at` source open (`A1/B1`, then `A2/B2`)
was:

| ordinary Parts | normal non-allocator A1 → B1 | normal non-allocator A2 → B2 |
| ---: | ---: | ---: |
| 256 | +0.87% / +0.74% / -0.63% / +0.85% | +4.50% / +4.22% / +4.03% / +4.41% |
| 2,048 | +3.41% / +3.81% / +3.27% / +3.48% | +3.83% / +3.91% / +3.17% / +3.73% |
| 16,384 | +1.40% / +0.26% / -0.43% / +1.32% | +1.50% / +2.19% / -0.69% / +1.30% |

Allocator-enabled latency is observational only. Its allocator vectors make
the retained-vector footprint evidence exact: 256 Parts remain at
`5,977` calls and `779,619` allocated bytes; 2,048 Parts add one call and
16,384 bytes; and 16,384 Parts add one call and 131,072 bytes. Deallocation
and reallocation vectors are unchanged. Every source-backed lookup size
remains at 48 allocation calls and 1,536 allocated bytes. Lookup source
counters remain zero reads, zero bytes, and zero ordinary payload bytes with
144 source-version calls; source-open counters are unchanged.

Eager normal-binary lookup timing is not decision-quality evidence: randomized `HashMap`
traversal produced large control/candidate drift (roughly 51–93% in sampled
directions), so no eager timing claim is made. The final source-open p50 and
mean overheads stay below 5%, but these normal-binary observations and the
allocator-enabled latency observations are not promoted to a latency claim.

## Verification and claim boundary

Focused case-fold index tests cover unsorted small catalogs, the deterministic
descending 2,048-Part boundary, canonical and alias hits, misses, preserved
iteration order, equivalent-name rejection, and managed fallback. The
managed cancellation/resource test confirms that a budget of one does not
pay for a retained index. `cargo test -p litchi-opc --lib` passed `282/282`
under explicit Rust 1.98.1, together with the exact rustfmt check and diff
check. Independent implementation review returned `SAFE`; independent test
validation returned pass.

`performance_claim: scoped`; `claim_authorized: true`. The authorized latency
claim is limited to normal, non-allocator unmanaged packages opened through
`SourceBackedPackage::from_read_at`, using the fixed 144-query vector on the
exact 2,048- and 16,384-Part corpora and final protocol. Validation-constructor
coverage is correctness-only. The exact allocator and retained-vector
footprint evidence and source-counter invariants above are reportable
evidence; allocator-enabled latency is observational only. No claim follows
for 256-Part lookup, source-open latency, eager lookup/open, managed packages,
mutable `OpcPackage`, typical OOXML/OPC packages, RSS, physical I/O,
decompression, cold-cache behavior, throughput, scaling, or general package
behavior. Means, tails, and the randomized eager timing are observations only.
Semantic correctness is limited to the stated corpus, fixed vector, and
focused tests.
