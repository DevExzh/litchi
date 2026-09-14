# Complex function performance evidence

Candidate 09 implements all 26 functions and passes the correctness gates
linked from [the parent report](../README.md). Performance evidence here is
scoped to the retained binaries, fixtures, and machine recorded in the raw
captures; remaining RSS flags prevent a broad no-regression claim.
[Host metadata](machine-09.json) identifies the AMD EPYC 9R45 KVM environment;
all measured children were pinned to CPU 6.

## New-family capture

[Release 09](release-09/receipt.json) retains 225 successful processes: 75
unique cases, each run three times on CPU 6 with two warmups and 31 measured
iterations. The corpus includes all 26 functions, aggregate scaling, reference
sequences, lazy branches, malformed input, and typed resource failures. Exact
harness source, command, compiler, 464-file source hashes, and retained ELF hash
are in `release-09/build.tar.gz`; exact runner and per-process stdout, stderr,
and `/usr/bin/time -v` records are in `release-09/capture.tar.gz`.

The pre-complex implementation refuses these functions. It is a capability
baseline, not a successful equal-work timing comparison. The earlier release
07 diagnostic has three failed text-limit oracles: it expected a formula error
where production returned the correct typed Memory refusal. That exact old
harness/capture remains under `diagnostics/release-07`; release 09 corrects the
oracle to Memory observed 32768, limit 32767.

The [analysis](release-09/analysis.json) reports per-case values. For example,
COMPLEX has a median 1,276 ns per evaluation, IMDIV 3,004 ns, and IMLN 1,823 ns
in this capture. These include the value VM, checksum and result drop, not
just arithmetic kernels. Each has zero retained result-budget memory; they
still incur evaluator allocations (9, 17, and 12 calls per evaluation,
respectively). No kernel-allocation claim should be read as an allocation-free
end-to-end evaluator claim. Numerical preflight is untimed; formula-error lanes
check the error kind category, while integration tests check exact codes.

Regenerate the summary without extracting scratch files:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-complex-functions/performance/analyze_release.py docs/report/spec-gap-validation-evidence/ods-formula-complex-functions/performance/release-09/capture.tar.gz
```

## Existing-value comparison

[Common 09](common-09/receipt.json) compares retained `value-matrix-06` and
`value-complex-09` ELFs using identical four-file value harness sources, limits,
fixtures, compiler and release flags. The pinned AB/BA/AB run contains 162
successful children and 81 pairs across 27 cases, with zero deterministic-field
mismatches. All binaries, harness bytes, runner bytes, and source receipts
remain unchanged across capture. Source archives identify the actual builds;
worktree HEAD is only context.

No per-case median p50 or p95 increase exceeds 5%. Six median RSS increases
cross the review threshold:

| Case | p50 | p95 | RSS |
|---|---:|---:|---:|
| reference-background-4096 | -1.59% | -4.16% | +6.70% |
| reference-cell | -0.57% | +0.37% | +9.25% |
| reference-empty | -1.71% | -1.38% | +7.07% |
| reference-matrix | +0.75% | +0.71% | +6.68% |
| reference-repeat-1 | +1.32% | +4.85% | +8.98% |
| reference-text | -0.38% | -1.98% | +5.19% |

Allocation counts, peak live allocation deltas, and retained-memory counters
match in this corpus. One requested/released-byte counter changes:
`matrix-lazy-aggregate-4096` increases from 3,146,824 to 3,146,840 bytes per
sample (+16 bytes). This is separate from the runner's deterministic-field
comparison, which does not establish equality of every memory counter.
Source inspection identifies a likely cause of the +16 bytes: adding the
24-byte Complex payload widens `DemandCacheValue`, and this case inserts one
cache entry for its AND branch. The unchanged 4,161 allocation calls and exact
one-entry byte delta support transient cache-entry overhead; private type
sizes were not measured, so this is an inference rather than layout proof.
The RSS observations do not identify an allocator or code cause. They remain
review items alongside the earlier matrix/common corpus flags.

The [full per-case analysis](common-09/analysis.json) is reproducible with:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-matrix-functions/performance/analyze_common.py docs/report/spec-gap-validation-evidence/ods-formula-complex-functions/performance/common-09/capture.tar.gz
```

All archive members were compared byte-for-byte with their loose inputs before
cleanup. Retained ELFs and the single disk-backed build workspace remain for
follow-up comparisons; no repository/build duplicates are placed in tmpfs.

## Cache layout follow-up

[Cache10 diagnostics](cache-10/receipt.json) tested a suffix-specific private
cache representation. Actual private-type probes against both source snapshots
report `Complex=24`, `DemandCacheValue=24`, and `DemandCacheEntry=32` bytes.
The earlier inferred 40-byte entry size was incorrect: Rust's enum niche
layout already keeps the original entry at 32 bytes. An empty cache reserves
two entries, so the earlier +16-byte delta is consistent with two entries
increasing by 8 bytes from the pre-Complex representation.

The 162-child AB/BA/AB comparison has no deterministic or instrumented memory
counter changes, and no per-case median p50/p95/RSS increase above 5%. This
window does not dismiss the earlier RSS flags; it compares a different,
non-beneficial prototype against candidate09. Root removed the prototype
because it saves no storage and adds checked reconstruction. The frozen
prototype source/gates remain in `../cache-10`, and exact paired capture,
build and instrumented layout-probe receipts are retained here. Public
cache suffix/read-reuse coverage and the 32-byte footprint guard remain useful.
