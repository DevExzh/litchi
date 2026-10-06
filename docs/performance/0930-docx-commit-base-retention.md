# 0930 — bounded DOCX commit base-snapshot retention

This batch adopts bounded reuse of a commit's validated base section inventory.
The primary generic ReadAt workflow falls from 4.487303 to 4.149882 ms pooled
p50 (7.52%). Six paired process ratios have median 1.081241 and 95% bootstrap
interval [1.079744, 1.083847], passing the frozen gate. All four routes improve;
no pooled total-latency tail or per-route process-RSS review threshold is crossed.

## Work and ownership

The 0929 native profile attributed 11.75% of observed publication period to
current main-document XML admission and section-inventory parsing. This batch
retains the validated base inventory already built by the
query, then shares it when commit publication observes identical XML and
semantic limits. That profile motivates the experiment; it does not predict a
normal-release speedup.

`Edit::with_max_retained_base_bytes` sets a separate optimization ceiling,
defaulting to 1 MiB. Changed source-backed commits keep the existing base
snapshot only when checked accounting fits. Zero, overflow, an exceeded cap,
or explicit `Commit::release_retained_base` selects the existing parse path.
`Commit::retained_base_bytes` exposes the logical retained weight. No-op commits
reuse their already-required target snapshot and report zero extra retention.
The retention ceiling is separate from semantic `section::Limits`, so changing
it cannot create a policy conflict. There is no policy-joining operation in
this path.

Accounting includes inventory arrays, nested vector/string capacities, source
spans and namespace bindings, tolerance-report payloads, Arc headers and one
logical Snapshot owner charge. XML is excluded because the patch already owns
the same allocation. Cloned commits share allocations. The always-present
inline optional-snapshot slot, allocator bookkeeping and physical RSS are not
the conditional retained-state measurement.

Publication still reads current XML, computes the full current artifact
fingerprint, checks source-version fences and execution state, and performs
the existing exact-byte, lineage, revision, artifact-fingerprint and limits
authorization before writing. Reuse requires exact XML and Limits equality;
otherwise existing admission, parsing and errors apply. The preservation writer,
output fingerprint, paragraph-cache trims and operation context are unchanged.
The publication's original snapshot remains ready for inverse publication;
work is not deferred into inverse validation.

| ADR | Constraint and implementation |
| --- | --- |
| 0003 | Immutable shared snapshots, atomic commits, exact source authorization and reversible publication remain intact. |
| 0005 | Existing allocations have an explicit finite retention ceiling, observable weight, release and recomputation fallback; full freshness checks remain. |
| 0031 | Budget-managed cache-backed section-layout sources retain their existing refusal; no second execution-budget system is introduced. |

## Evidence scope

The 24-child unchanged-binary precheck is descriptive only. Formal comparisons
must use fresh interleaved baseline/candidate children, six pairs per route,
five warmups and fifty measured iterations per process on CPU 12. The primary
route is generic ReadAt; owned vector, owned reader and short range are separate
regression controls. The gate is median paired baseline-p50/candidate-p50 at
least 1.05 with lower 95% bootstrap bound above one (10,000 pair resamples,
seed 930). Pooled tail and setup-inclusive maximum-RSS regressions above 5%
are review flags. This tests one generated warm in-memory DOCX workflow, not
cold storage, physical I/O or an actual remote service.

The original precheck plan accidentally reversed the future formal ratio. Its
bytes remain retained, and a prospective correction fixed the ratio before any
candidate timing run. The precheck runner also printed a misleading failure
line while catching `SystemExit(0)`; its original tool exit was zero and all
24 reports passed. Neither issue is hidden by repeating those children.

## Build and quality history

The first candidate release build failed with ENOSPC, including its attempt to
write the terminal receipt. The original tool exit and compiler log were saved
on the main filesystem before failed build objects were removed. A new release
build uses a target outside tmpfs. No candidate timing workload ran from the
failed build.

The root workspace lock prevented the initial `--locked` quality command from
compiling. Quality runs use an isolated checkout and the pinned canonical 0922
lock; the shared root lock is untouched. Performance uses the unchanged
standalone benchmark lock. Two real-document unit-test fixtures hit the
existing markup-compatibility refusal before the new retention assertions.
The final test constructs a minimal OPC package; no refusal rule was weakened.
All unsuccessful commands remain in the evidence record.

Final quality checks pass: 1,921 all-target tests across 73 suites (one ignored),
78 doctests (31 ignored), all 24 source-backed section integration tests,
warning-denied Clippy, warning-denied rustdoc, and formatting for all five
changed DOCX files. The dependency-boundary checker passes with 65 packages,
244 internal declarations and 11 existing explicit debt items. Workspace-wide
formatting reports unrelated differences; this batch does not change them.
The reduced quality receipt binds all command attempts, source and fixture
hashes, and lock provenance. Root independently verifies 50 referenced
receipt, output-log and final-source hashes.

The successful normal candidate ELF is 70,197,424 bytes, SHA-256
`ae9d7c303496a30ae2850e397076093322adb6fe2be4a598ac446b206bd88eff`.
It uses Rust/Cargo 1.95.0 and the same standalone benchmark lock and normal
release command as the retained baseline ELF (70,187,352 bytes, SHA-256
`7138c6c458dc13f1677deba3f413120ffc1af155d07848a243a43ebcf335feb7`).
The exact v2 build snapshot is reconstructed and matches all 6,276 source
manifest entries. The later minimal-fixture correction changes only the
`#[cfg(test)]` suffix of layout.rs; the 83,018-byte runtime prefix is identical,
SHA-256 `951cd80a131c00d2a4a42d2ea061e6444becfff23e1c84f03391be9a641f6e28`.
The source review and separate supersession receipt bind both versions. No
extra release build is claimed for that test-only correction.

## Native result and limits

All 48 scheduled children pass, with 2,400 measured samples and no retry or
selective recapture. The original outer tool handle returns zero; its receipt
is saved before replay and reduction. Original and moved packets replay with
identical metrics. The [independent audit](results/change-0930/review/independent-formal-audit-v1.json)
checks all 48 reports and 576 semantic gates; its copied reduction matches
every route and paired statistic. Each child retains exact input/output hashes, all twelve
correctness checks, provider counters, six phase arrays, process status and
setup-inclusive maximum RSS. Source/output sizes remain 2,101,427/2,101,446 bytes.

| Route | Baseline pooled p50 (ms) | Candidate pooled p50 (ms) | Paired ratio median | 95% bootstrap interval | Maximum RSS change |
| --- | ---: | ---: | ---: | --- | ---: |
| Generic ReadAt (primary) | 4.487303 | 4.149882 | 1.081241 | [1.079744, 1.083847] | +3.24% |
| Owned vector | 3.717930 | 3.379289 | 1.100994 | [1.097352, 1.179917] | +3.54% |
| Owned reader | 3.843231 | 3.513425 | 1.093107 | [1.052091, 1.130305] | +1.32% |
| Short range | 4.533633 | 4.192777 | 1.080853 | [1.079999, 1.083198] | +2.85% |

The primary publication p50 falls from 2.671086 to 2.337030 ms (12.51%), while
query p50 remains approximately 1.33 ms and commit approximately 0.35 ms.
Generic total p95/p99 fall from 4.515988/4.544588 to 4.179217/4.196837 ms.
Across routes, pooled total p95/p99 improve; maximum process RSS increases
1.32–3.54%, below the 5% review threshold. Median process RSS also increases
2.40–4.31%. RSS includes setup, verification and allocator effects and cannot
be equated with the logical retained-base gauge or operation peak-live bytes.
No allocation-count or zero-memory-cost claim is made.

Owned-reader process medians vary more than the generic route, and one
owned-vector candidate process is substantially faster than its peers. Every
child remains included; the primary result does not pool these control routes.
These six-pair intervals describe this warm, generated fixture on the recorded
machine. They do not establish behavior across large inventories, many live
commits, cold storage, or other Office formats.

## Reproduction and retained evidence

The [authorized manifest](results/change-0930/receipts/formal-manifest-authorized.json)
binds the source closures, normal binaries, toolchain, route order, exact
child environment, workload arguments, validator, runner and corrected gate.
The [quality summary](results/change-0930/new-quality.json) and source/runner
reviews under `results/change-0930/review/` bind the separate correctness checks.
The [reduction](results/change-0930/formal-reduction-final.json) retains all
process-pair medians, intervals, pooled phase statistics and RSS vectors.
The [capture/replay receipt](results/change-0930/receipts/formal-capture-replay-receipt-final.json)
records original and moved replay evidence. The runner creates receipts
exclusively, validates with its packet-local validator, uses a finite explicit
child environment and verifies process-group closure. Scripts stayed frozen
through capture. The shared machine was not guaranteed globally quiescent.

Offline replay, without another document workload:

```sh
python3 -B docs/performance/results/change-0930/comparison-formal-runner-v1.py replay --packet /tmp/litchi-performance-0930/formal-packet-authorized
```

Replay requires the externally retained source/build evidence and binaries
named by the manifest. The archived packet is
`/home/zhuhe/litchi-performance-0930/formal-packet-authorized.tar`, 6,062,080 bytes,
SHA-256 `2630e4d2f03c6434579bcb19b5f98fcfae8ecf3746fcad48679088272c9beb85`.
Its [archive receipt](results/change-0930/receipts/external-packet-archive.json)
records the regeneration command. All large evidence and executables remain
outside Git; the repository contains reduced evidence only.

Cleanup removes 5,492,353,322 logical bytes of disposable build objects and
focused-test scratch trees, including an unexecuted later build wrapper.
Physical reclaimed space is not measured. The
[cleanup summary](results/change-0930/receipts/cleanup-summary.json) binds the
removed-path receipt and preserved artifacts. Shared build caches and unrelated
work remain untouched.
