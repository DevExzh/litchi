# 0929 — native phase ancestry with frame pointers

This batch builds a separate diagnostic executable to recover the query and
publication ancestry missing from 0928. It changes no production or benchmark
source. Its elapsed times are not compared with normal release measurements.

## Build and observation boundary

The source is the retained 0925 v4 closure. Its non-test harness prefix remains
byte-identical to the current checkout; the current harness adds only the
0927 diagnostic tests. Rust/Cargo 1.95.0 builds with
`-C force-frame-pointers=yes -C debuginfo=1`. The first successful release build
stripped debug sections. A successful final-bin `cargo rustc` relink with
`-- -C strip=none` preserves the dependency objects' debug information.
Both build results are retained; only the corrected executable was profiled.

The corrected ELF is 529,059,232 bytes, SHA-256
`c195a1d94acb7b240ca86ad32c13a9fd8b4c7b3affe085122970fe0f75c2169e`.
Independent disassembly and source-line checks verify its frame-pointer
prologue and exact sample-loop calls: query at `0xf49831`, returning to
`0xf49837`; publication at `0xf49cb8`, returning to `0xf49cbd`.
Expected-output setup, output verification and inverse publication have
different call sites. These addresses apply only to this ELF.

One generic ReadAt document workload runs 1,000 iterations with zero warmups,
on CPU 12, under `perf record -F 997 -e cycles:u --call-graph fp`, with a
120-second deadline. The original handle returns zero and its terminal receipt
is persisted before decoding. Perf also inherits metadata subprocesses.
The profile includes corpus construction, preparation and correctness checks
outside the timed phases.

## Native attribution

Perf records 15,966 events, 63,813,450,383 total period units and zero reported
lost samples. The leaf and physical-stack exports join exactly, with no missing
or duplicate keys; all first-frame instruction pointers and DSOs agree.
Every sample has at least one printed frame. This does not prove every chain
reaches its root: unresolved frames remain, and system libraries need not
preserve frame pointers.

The capture-specific PIE load bias is `0x5b8705390000`, derived from the
recorded executable mapping and ELF load segment. Only frames belonging to
the workload ELF are normalized against it. Exact caller addresses establish
these mutually exclusive buckets, with matching query/publication callee
frames also required:

| Attribution | Samples | Period units | Whole-profile share |
| --- | ---: | ---: | ---: |
| Query | 1,504 | 5,953,322,348 | 9.33% |
| Publication | 2,804 | 11,394,220,603 | 17.86% |
| Verification and inverse checks | 8,247 | 32,943,877,871 | 51.63% |
| Expected oracle and metadata subprocesses | 21 | 72,783,157 | 0.11% |
| Unclassified | 3,390 | 13,449,246,404 | 21.08% |

The decoder maps no samples to separate edit, commit or setup buckets; this
means no mapped observations, not zero work. Unknown weight remains in the
full denominator. Fractions within an attributed phase describe only that
observed subset, not all elapsed time in the phase.

Hardware-assisted SHA-256 compression contributes 4,228,981,709 period units
within query (71.04%) and 8,570,643,406 within publication (75.22%). Publication
SHA samples have either a source-fingerprint ancestor (4,213,757,344 units) or
an output `FingerprintingWriter::write` ancestor (4,356,886,062); none belong
to both or neither. Their near-even split is consistent with the distinct
source and output hash responsibilities, and does not authorize skipping one.

The inclusive union of `Inventory::parse_with_limits` and
`ensure_source_section_inventory_xml` ancestry covers 1,338,536,419 period
units, 11.75% of attributed publication: 7.41% parsing and 4.34% admission
checking, with no overlap in this capture. These are inclusive samples, not
self costs to add to the leaf table. One sample has the publication return
address but commit-cleanup ancestry and no publication callee. It is classified
as unknown; a caller address alone is insufficient evidence for that sample.

The initial default stack export expanded inline frames and emitted eleven
addr2line diagnostics. It is retained. Explicit `perf script --no-inline`
produces the canonical physical-frame export with empty stderr, using the
same capture. No document workload was repeated for decoder problems.

## Correctness and scope

The report retains the same deterministic 2,101,427-byte input and
2,101,446-byte output as earlier batches. All twelve reported correctness
checks pass. This is diagnostic CPU evidence, not a replacement for the
unqualified 0925 timing cohort or a before/after speedup measurement.

No cache, source-freshness shortcut or production optimization is adopted.
The next candidate is to retain a commit's validated base layout snapshot and
reuse it during publication. The measured repeated parsing/admission work
justifies a bounded experiment, not adoption: capture normal-build before/after
latency and memory, preserve exact source XML, full fingerprint, version,
lineage, limits, cancellation and publication checks, and test mismatch/refusal
and inverse parity. Retaining the base inventory adds commit-held state; it
must have explicit resource accounting. This diagnostic build supplies no
speedup bound for the normal release build.

## Reproduction and retained artifacts

The build and profile receipts under `results/change-0929/receipts/` record
the exact commands, source closure, locks, toolchain, executable identities
and original terminal status. Large evidence stays outside the repository:

- `/tmp/litchi-performance-0929/capture-fp/perf.data`: 2,855,464 bytes,
  SHA-256 `9a4aa1051bfdbad3c148948cd25c5164c07181fa2d45169847c4c3859619c93f`.
- The corrected executable remains at
  `/tmp/litchi-performance-0929/target-fp/release/litchi-perf-baseline`.

The source and output hashes, report arrays, artifact sizes and regeneration
commands are retained in the receipts. Runtime source is unchanged, so this
batch validates the diagnostic report and offline analysis rather than
rerunning Cargo correctness suites.

The committed decoder is `results/change-0929/analysis/decode_fp_phase_ancestry_v1.py`;
its full command is in `analysis/fp-phase-decoder-receipt-v3.json`. It reads the
two retained exports and the static map/load-bias receipt, without launching a
document workload. The publication addendum and report validator are also
committed; those two scripts intentionally use this capture's external paths.
Root replay of the committed decoder returns zero and reproduces the full
summary exactly. The report validator returns zero, and independent phase
counting agrees on query/publication samples, weights and SHA leaf totals.
Wrong-binary and wrong-load-bias decoder controls return failure.

Cleanup removes approximately 4.05 GB of disposable build objects and
regenerable analysis intermediates. Both executable identities, the original
capture, adverse decoder outputs and final supporting evidence remain
external. No compiled executable or evidence file of 1 MiB or more is committed.
