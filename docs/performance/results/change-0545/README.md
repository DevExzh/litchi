# 0545: exact-bound scanner diagnostic rejected

The isolated chunked scanner is **rejected as an unconditional replacement**.
Production is unchanged. Dense input improves, but sparse XML takes 42.98–44.18x
the baseline p50 time, and the early-rejection input regresses 22.61–22.73%.
This is attribution evidence for the active OOXML optimization, not a retained
runtime or end-to-end improvement. OLE2/OOXML remain first; ODF is deferred until
that goal completes.

## Frozen experiment

`freeze.json` predates all four quality commands and 28 native processes. It binds
sources, dependency lock, fixtures, ADR hashes, environment, CPU 2, 10 warmups,
100 single-call samples per process, seven fixtures and A1/B1/B2/A2 order.
`capture.py` records exact commands, stdout/stderr hashes, timestamps, exits,
freeze hash and measured executable hash. Runner/analysis scripts are sealed
with the final evidence, rather than claimed as pre-capture frozen inputs.
Both functions reside in one release executable, both have `inline(never)`,
and selection is by runtime function pointer. Fixture loading, direct NsReader
oracle, warmup, result assertions, storage and JSON writing occur outside timed
calls. All 2,800 samples remain; no rerun or filtering occurred. The initial fmt
and offline lock-resolution setup precede the frozen quality commands.

The baseline is the exact 0544 two-delimiter-scan helper, renamed and given the
same call boundary as the candidate. The candidate sums the **same** full bound:
EOF + initial nonmarker + count(`<` or `&`) + count(`>` or `;` followed by a
nonmarker). It partitions adjacent pairs into 4,096-byte chunks and counts the
last marker separately. Subtotals are at most 8,192 and total additions are
checked. It does not introduce unsafe code, allocation, dependencies in the
workspace, or a new event limit. The direct parser validates event-count bounds;
the helper is not an XML validator.

Fixtures 160/164/256 are exact worksheet bytes extracted from sealed 0544 cap
archives. Numeric 96/128 are synthetic diagnostic XML, **not** the original
primary workflow fixtures. Sparse and early-reject fixtures are lexical stress
cases. `generate.py` reconstructs all seven; `fixtures.json` pins provenance and
hashes. Every process checks full-bound eligibility and direct NsReader event
count. The 128 diagnostic remains eligible (82,181), unlike the earlier coarse
bound proposal; exact equivalence also preserves eligibility for original inputs.

## Results and decision

Percentages below are candidate p50 change versus baseline, two matched repeats.
All mean/p95/p99 values, repeat drift and individual >5% reviews are in
`summary.json`, reproducible using `analyze.py`.

| Fixture | Repeat 1 | Repeat 2 |
| --- | ---: | ---: |
| numeric 96 | -43.616% | -44.859% |
| numeric 128 | -43.159% | -42.669% |
| cap 160 | -39.974% | -41.789% |
| cap 164 | -38.971% | -37.230% |
| cap 256 | -11.006% | -10.245% |
| sparse 1 MiB text | +4317.818% | +4198.277% |
| comment prefix, early rejection | +22.730% | +22.606% |

The 18 individual reviews include 16 adverse comparisons and two repeat-drift
flags. Two process repeats and complete sample distributions expose variability;
there is no formal population confidence interval. Do not infer workflow,
cold-cache, allocation, RSS, hardware-counter, instruction-count or concurrency
results. Neither function allocates by source inspection; allocator/RSS effects
of integration were not measured. There is no claim that existing late-validator
regressions from 0544 are fixed.

`assembly.stdout.gz` losslessly preserves complete disassembly of the measured binary. Candidate
symbol `xlsx_scanner_diagnostic::chunked::chunked` at 0x23f30 uses scalar byte
loads and conditional jumps in its inner loop (0x24010–0x24066), not a vector
reduction. Dense gains alone do not justify replacing sparse delimiter skipping.
The source comment describes intended compiler opportunity, not achieved SIMD.

Next: investigate an exact-count implementation that preserves sparse skipping
and cheap early cap exits. Non-short-circuit boolean reduction is an **unmeasured
idea**, not a selected or validated fix; vectorization alone would not establish
acceptable sparse behavior. Recheck actual generated assembly and the same
stress cases before another full 0544-style correctness/workflow/allocation/
refusal/cap campaign. The rejected coarse bound remains unsuitable because it
excludes the primary 128 shape. Fresh commit-path attribution remains subsequent
work after a shared-traversal revision earns admission.

## Validation and reproduction

Four frozen diagnostic checks pass: fmt, two oracle tests, warning-denied Clippy
for all diagnostic targets, and release build. The tests cover chunk boundaries,
exact cap edges, XML events/references/malformed input, BOM, and 512 deterministic
arbitrary-byte cases. All native-process assertions pass. No workspace runtime
source changed; repeating workspace suites, rustdoc, Office fixtures, fuzzing,
Miri or concurrency tests was not applicable to this isolated evidence batch.

From the repository root, create a fresh isolated directory with `Cargo.toml`,
`Cargo.lock` and `src/{main,baseline,chunked}.rs` copied from this bundle. Generate
fixtures using `python3 -B docs/performance/results/change-0545/generate.py DIR`.
Run the exact Cargo commands in quality receipts with the manifest path adjusted
to that directory, then invoke the release binary as `taskset -c 2 BINARY
baseline|chunked FIXTURE OUTPUT.json` in frozen order. `capture.py` documents the
original absolute paths and refuses to overwrite evidence. Adjust its output
and scratch roots for a fresh campaign; do not replace this sealed evidence.
Run `python3 -B docs/performance/results/change-0545/verify.py` to verify this
bundle and regenerate fixture hashes without building Rust.

The initial staged whitespace check flagged spaces emitted by objdump and the test runner’s final blank line. Its
output is archived in `initial-diff-check.stdout.gz`; disassembly and test stdout are preserved
losslessly as gzip and the verifier checks its original receipt hash after
decompression. Captured bytes were not normalized. The final staged check passes.
