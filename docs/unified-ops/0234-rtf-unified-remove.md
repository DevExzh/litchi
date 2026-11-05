# 0234 — Unified RTF paragraph removal

This milestone connects retained RTF paragraph removal to `File::remove`,
explicit transactions, atomic/best-effort batches and process-local history.
Scoped query targets resolve against one immutable source snapshot. The native
batch removes the selected original paragraph positions atomically, with no
sequence of facade-owned reparses or independently published intermediate files.

The native editor preflights ordered selectors and operation bounds before
staging. Publication uses validated ordinary paragraph source ranges and merges
adjacent removal ranges. The facade additionally bounds expanded targets, staged
retained text/action bytes and snapshot history. Native source policy, protection
and exact process-local inverse replay remain authoritative. Non-paragraph
removal and unsupported source closures return typed refusals.

## Verification

The complete native RTF suite passes 1,474 tests/doctests; the full-feature
facade suite passes 1,650; the RTF-only facade suite passes 515. Eleven new
unified integration tests cover scalar native byte parity, scoped fanout,
original-coordinate removal, empty survivors, trailing boundaries, control-word
delimiters, sequential mixed edits, exact history, batch rollback/best effort,
typed limits/protection, origin/save/reopen, and explicit native repair policy.
Nine native batch tests cover admission, source boundaries and replay contracts.

All-target warning-denied Clippy passes for native/facade all-features, the
RTF-only facade, and the standalone benchmark harness. Warning-denied rustdoc,
no-feature compilation, scoped formatting and diff checks pass. Crate boundaries
pass for 66 packages; 36 Python repository-tooling tests pass; the generated
OfficeCLI delta is unchanged. No new Python binding surface is claimed.

Source-range qualification found and fixed three preservation details: an empty
survivor requires its terminating paragraph boundary; deleting an adjacent
boundary must not form controls such as `\parThird` or absorb numeric/sign/space
prefixes; and the final `\par` does not add another native semantic paragraph.
The established scalar lifecycle path remains available for decoded and
explicitly repaired sources without literal ranges. Multi-paragraph removal
requires the narrower literal ASCII/default-formatting closure. Its native patch
has exact process-local inverse/replay; durable multi-removal serialization is
explicitly refused pending source-bound preconditions, rather than exporting
sequential operations with incorrect original coordinates.

## Performance

The standalone `tools/perf-baseline/unified-rtf-remove` harness compares native
and unified single-paragraph and scoped-query removal on 10- and 1,000-paragraph
documents. It uses warmed immutable owners, prepared selectors, alternating ABBA
samples, exact byte comparisons and semantic reopen checks outside timed regions.
Three final trials ran on one pinned CPU on a shared host, with two warm-ups
and nine ABBA rounds each (18 samples per route/case), over 1, 10, 100, 1,000 and
10,000 paragraphs. Every sample passed exact native/unified byte and semantic
reopen oracles. Receipts include source/tree hashes, binary size/SHA-256, compiler,
protocol and raw timings: [trial 1](results/0234-rtf-remove/trial-1.json),
[trial 2](results/0234-rtf-remove/trial-2.json),
[trial 3](results/0234-rtf-remove/trial-3.json).

Ranges of per-trial medians (microseconds):

| Operation | Source paragraphs | Native µs | Unified µs | Unified/native |
|---|---:|---:|---:|---:|
| single_path | 1 | 2.430–2.490 | 2.770–2.790 | 1.112–1.144× |
| query_all | 1 | 2.360–2.420 | 3.450–3.510 | 1.426–1.470× |
| single_path | 10 | 7.100–7.190 | 7.500–7.610 | 1.050–1.072× |
| query_all | 10 | 6.790–6.810 | 8.530–8.550 | 1.256–1.256× |
| single_path | 100 | 45.351–45.660 | 46.290–46.790 | 1.021–1.032× |
| query_all | 100 | 40.910–41.531 | 48.341–48.490 | 1.168–1.182× |
| single_path | 1,000 | 539.752–546.992 | 541.682–545.332 | 0.990–1.008× |
| query_all | 1,000 | 373.882–377.982 | 434.842–439.771 | 1.163–1.170× |
| single_path | 10,000 | 6473.527–6548.219 | 6485.347–6515.607 | 0.993–1.007× |
| query_all | 10,000 | 3824.848–3849.397 | 4387.760–4416.889 | 1.147–1.152× |

The facade now charges removed payload from already-resolved native paragraph
nodes and uses the scalar native entry for one target, avoiding an additional
story walk and native source-map build. Single-target removal is near native
cost on the larger documents; the one-paragraph case remains about 11–14% above
native (roughly 0.3 µs), exceeding the 10% small-write comparison reference.
Reducing fixed dispatch/owner-allocation cost is a performance follow-up.

The 1,000-target fanout is within 2× native, but this is one scoped selector,
not 1,000 separate batch commands. Repeated-command coalescing, allocation
qualification and the complete performance acceptance gate remain open.
These shared-host measurements support the stated cases only.

Reproduce using the harness README command with
`--counts 1,10,100,1000,10000 --warmup 2 --rounds 9`, three times, pinning an
available CPU. The measured source is `95984af5d8`. Validation/builds used external
copied locks and target directories; all three repository lockfile hashes remain
unchanged. Both owned agent worktrees/branches were removed. Owned 0234 scratch
and build output are removed after qualification; only small JSON receipts and
source are committed, never the benchmark binary.

## Remaining work

RTF add/move/swap, richer native property and node projections, transaction
coalescing, the remaining glossary migrations and other all-format requirements
remain active. This milestone does not establish full-goal completion.
