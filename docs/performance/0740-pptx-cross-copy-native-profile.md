# 0740 — native attribution of owned PPTX cross-copy planning

Generated-entry Deflate is a concrete next investigation target for the
media-rich corpus: it accounts for **78.79–79.54%** of strictly attributed
planning sample period in three fresh diagnostic profiles (median 78.94%).
The plain corpus gives 1.77–2.44% (median 2.12%). Production and Rust harness
source are unchanged. This is sampled cycle attribution in a frame-pointer
build, not an optimization result, an ordinary-release time fraction, a
removable-cost estimate, or a speedup prediction.

## Evidence and qualification

The source is `04210d2c78`; all 7,281 tracked crate/harness source and manifest
inputs match 0739. Four supplementary workspace inputs and 34 goal/checklist/
ADR hashes are bound. Both binaries were built release/offline/locked with two
Cargo jobs. The separate diagnostic binary adds exactly
`-C force-frame-pointers=yes -C force-unwind-tables=yes`; it uses the same
existing generated corpora and full report projection as 0739. All emitted
reports pass that logical same-stack oracle. This is neither a raw ZIP
physical-metadata oracle nor native Office validation.

Ordinary release profiling failed callchain qualification. The first capture
using perf's compression is corrupt and excluded. Uncompressed DWARF captures
at 16,384 and 65,528 stack bytes decode, but do not reliably connect planning
to the lifecycle root. The large lifecycle stack frame supports a bounded
stack-window explanation without proving the unique cause of every unwind
failure. All attempts are retained, including the corrupt capture. See the
[method review](results/change-0740/method-review.md).

The frame-pointer pilot passes both corpus projections in native and recorded
runs and exposes complete planning chains. Only then were classifier and
formal commands frozen. The six formal processes run serially on CPU 12,
alternating case order across three repeats, with 100 plain or 10 media-rich
samples and zero warmups. Recording uses `cycles:u`, 997 Hz and `fp,127`.
Startup, corpus construction and post-clock oracle work are present in the
whole-child capture and explicitly excluded from strict planning. Recorded
elapsed arrays are retained but are not used for latency or observer-overhead
claims. No timing equivalence with the ordinary build is asserted.

## Frozen classifier and denominators

A strict planning chain must contain, in leaf-to-root order, `prepare`, the
out-of-line `plan_cross_slide_copy_for_slides` helper, and
`run_pptx_cross_copy_lifecycle`, with neither apply marker. Unknown, empty,
depth-limit and conflicting chains remain ambiguous. Every sample period is
assigned once to planning, rooted apply, rooted other, unrooted work, or
ambiguous. Within strict planning, the generated-Deflate bucket also requires
`build_candidate`, `PackageWriter::write_to_stream`, `generated_entry`, and
`zlib_rs::deflate::deflate`. Its residual buckets are candidate writer other,
candidate other, and plan other. These do not pretend to recover every
internal source phase or identify individual ZIP members.

| Formal process | Corpus | All samples | Strict planning samples | Generated Deflate / strict planning period |
|---|---|---:|---:|---:|
| 0 | plain | 1,374 | 331 | 2.1151% |
| 1 | media-rich | 7,732 | 2,963 | 78.7909% |
| 2 | media-rich | 7,809 | 3,040 | 79.5407% |
| 3 | plain | 1,389 | 339 | 1.7729% |
| 4 | plain | 1,382 | 328 | 2.4402% |
| 5 | media-rich | 7,754 | 2,969 | 78.9364% |

The denominator is retained strict-planning **period**, not sample count or
whole-child period. Per-process excluded periods and all exclusive residuals
are in [analysis.json](results/change-0740/analysis.json). Inclusive symbol
periods are labeled overlapping and must not be added together. Unknown-chain
exclusion can bias attribution; the narrow spread across three processes is
only descriptive repeatability, not a confidence interval or proof of
unbiased sampling. No fraction is multiplied by 0739's planning wall-time share.

## Source-backed next candidate

The current owned candidate transfers decoded `BlobPart` data; newly appended
members reach generated-entry Deflate. The source-backed copy path already
has an OPC-owned authorized precompressed transfer mechanism. That contrast
motivates investigating bounded reuse for copied binary image leaves. The
profile locates generated compression under planning; it does not prove which
members dominate or that all such work can be removed.

Any implementation must preserve source freshness, decoded-byte identity,
CRC/size/content-type proofs, resource accounting and fresh destination ZIP
framing. It must avoid retaining an unbounded foreign archive. Durable patches
rebuild and compare target physical identity, so changing compressed bytes
requires an explicit compatibility or fallback design. Reversible patches,
refusal behavior and all publication proofs remain mandatory. No speculative
production wiring was retained. [Source review](results/change-0740/source-review.md).

## Replay and limitations

From the repository root:

```sh
python3 -B docs/performance/results/change-0740/artifact-seal.py --check
python3 -B docs/performance/results/change-0740/analyze.py
python3 -B docs/performance/results/change-0740/audit.py
python3 -B docs/performance/results/change-0740/audit-root.py
```

`freeze.json` binds the classifier and formal runner before capture. Receipts
bind exact argv, process times, binary identity, reports, raw perf bytes,
symbolized stacks and stderr. `raw-archives.json` binds lossless offline gzip
archives to original raw-byte hashes; the rejected perf-compressed stream
remains corrupt after lossless extraction. Offline gzip does not repair it.
The numeric replay uses retained symbolized stacks and needs no executable.
Fresh raw resymbolization requires the matching ELF, whose hash and selected
symbol evidence are retained; owned build trees and their perf ELF caches are removed after validation.
A fresh source rebuild is not claimed to reproduce that ELF byte for byte.

Independent replay and nine malformed/context controls validate the classification.
A supplementary replay verifies all 13 raw archives, percentages and 13 controls;
three documentation/coverage structural gates apply. No production test or
performance coverage promotion is claimed. Instructions, IPC, cache misses,
allocation, RSS, I/O, recompressed-byte counts, scaling and causal before/after
results were not measured here. The broader goal remains open.
