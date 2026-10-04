# 0905 — bounded XML attribute-summary candidate

Status: accepted after both primary native gates, the micro guard, quality
checks, independent review and retained replay passed. Both owned temporary
roots are removed, and post-cleanup replay passes. Integration into the target
branch remains pending while an unrelated shared-checkout cherry-pick is active.
The broader Office performance program is not complete.

This experiment starts from commit `4f39f5a88d`, the 0904 XLSX attribute-census
qualification. It tests one production-file change in
`crates/xml-minifier/src/audit.rs`. The candidate records a bounded
`AttributeSummary` during the existing complete `check_start` pass. A tag with
no attributes returns the inherited whitespace state without a second
attribute iterator; a tag with one raw attribute other than `xml:space` charges
the existing aggregate attribute budget once and returns the inherited state.
Multiple attributes and every exact raw `xml:space` spelling retain the checked
iterator, duplicate handling, value validation and per-element accounting.
The first pass, namespace checks, limits, error ordering, stream offsets and
replacement-window behavior remain in place. There is no public API change.

The captures, replay and final source check were completed against the frozen
`4f39f5a88d` source scope. Concurrent later iWork integration commits
`0e1f7a87a1` and `59f75c7f79` occurred after that evidence was made and are
outside this measurement. The current shared worktree therefore must not be
read as a byte-for-byte replacement for the frozen measured source archive.

The isolated source inventory contains 10,773 files, matching the retained
0904 harness inventory. The candidate manifest differs at exactly
`crates/xml-minifier/src/audit.rs`:

| source | bytes | SHA-256 |
| --- | ---: | --- |
| baseline `baseline-audit.rs` | 124,782 | `91424b57dfcecaf6cb204429e6d779a0c81a8b55c6546d616c694f4ea2455ac0` |
| candidate `candidate-audit.rs` | 126,817 | `eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7` |

The source and build custody records are retained in the
[0905 evidence packet](results/change-0905/README.md), including the exact
source manifests and the candidate patch.

## Semantic evidence

The differential oracle covers 11,520 generated cases and emits 125,952 raw
records, totaling 74,265,040 logical bytes. Baseline and candidate output are
byte-identical. The comparison records logical SHA-256
`567535be0da4b7165e65ed9d923fb1a9f15b39ffb5451133a876a6c9d22b0bc3`.
The retained [parity comparison](results/change-0905/parity-comparison.json)
is the machine-readable result.

The corpus spans six attribute cardinalities (0, 1, 2, 31, 32 and 33),
128 lexical styles, 15 resource profiles, slice/source,
authored and compact policies, two reader policies, three chunkings, full and
window replacement paths, malformed inputs and limit failures. The reusable
oracle and its method are retained in
[parity-method.md](results/change-0905/parity-method.md), the retained
[parity generator](results/change-0905/parity-example-v2.rs), and its final
source binding in [examples-v3.json](results/change-0905/examples-v3.json).

The independent [parity review](results/change-0905/parity-review.md) records
two corpus qualifications. The style formula
selects values 0 through 127, so the `style & 0x80 != 0` XML-declaration branch
described by the method is unreachable in this finite corpus. Also, the
diagnostic record hash uses the implementation's custom FNV-like multiplier
`0x1000000001b3`; it is not the standard FNV-1a prime. The raw JSONL equality,
gzip custody and logical SHA-256 remain the authoritative parity checks. These
limitations do not change the byte-identical result, but they narrow what the
oracle corpus demonstrates.

The source review found no counterexample within the candidate boundary. It
specifically checked malformed tails, duplicate names, `xml:space` values,
namespace declarations, aggregate and per-element limits, and slice,
streaming and replacement-window call sites. This is a source qualification,
not a substitute for the differential oracle or performance gates; see
[source-review.md](results/change-0905/source-review.md).

## Quality and custody

The retained quality matrix reports:

| check | passed | failed | ignored |
| --- | ---: | ---: | ---: |
| XML tests, baseline | 134 | 0 | 1 |
| XML tests, candidate | 134 | 0 | 1 |
| OPC/XLSX/DOCX/PPTX tests, candidate | 6,296 | 0 | 3 |
| documentation tests, candidate | 93 | 0 | 33 |

Candidate formatting, warning-denied Clippy, warning-denied rustdoc and the
crate-boundary check passed. The first baseline XML invocation retained a
compile failure caused by an unnecessary `std::io::Read` qualification in the
new standalone oracle; the corrected v2 invocation passed and no oracle was
run from the failed build. An initial bootstrap attempt found the isolated
workspace `Cargo.lock` missing before launching Cargo; the retained 0902
quality lock was copied into the isolated source root, the helper was not
changed, and all subsequent child commands recorded unchanged inputs. These
are retained custody events, not discarded measurements.

The standalone parity and probe examples are excluded from the 10,773-file
benchmark inventory and are bound separately by
[examples-v3.json](results/change-0905/examples-v3.json). The packet's
pre-execution review froze the main and micro helper hashes, source bindings,
role bindings and write-once artifact rules before performance execution.

## Performance protocol and result

The planned main comparison uses four fresh-process cases: medium and
dense-sparse XLSX one-percent edits as the two primaries, medium XLSX one-edit
as a control, and medium DOCX semantic one-edit as a control. Qualification is
eight children with two samples and no warmup. Native measurement is six
alternating paired blocks, 100 samples and ten warmups per child. The separate
allocator run uses two repeats with three samples and two warmups; it is
diagnostic and does not enter the native latency estimate.

Adoption requires both primary paired p50 median ratios to be at most 0.98 and
both 10,000-draw paired-median bootstrap upper 95% bounds to be below 1.0,
with seed 905009. p50, p95 and p99 regressions above 1.05 and role/case p50
spreads above 1.20 remain review flags. The medium and dense-sparse primary
workloads are required to pass independently; controls cannot substitute for
either primary.

The frozen micro guard exercises 12 fresh children over 36 jobs and 30 samples
per job, for 12,960 timed batch samples, with six alternating paired blocks and
seed 905019. Its hard rejection condition is a paired p50 median above 1.05
together with a 95% lower bound above 1.0. The main and micro gates are both
required before adoption.

The completed [main reduction](results/change-0905/analysis.json) has 48 native
children and 16 allocator children. Both primary native rows pass the frozen
gate, with no main review flags:

| case | paired p50 median ratio | bootstrap 95% interval |
| --- | ---: | ---: |
| XLSX one-percent, medium | 0.9543881780 | 0.9508342443–0.9584711179 |
| XLSX one-percent, dense-sparse | 0.9477553767 | 0.9350215878–0.9742145402 |

The XLSX one-edit and DOCX controls are descriptive at 0.9795903412 and
0.9700651277 respectively; their upper intervals are 0.9903041916 and
0.9881021201. Controls do not enter the primary adoption decision. The
publication phase paired ratios are 0.9051543375 for medium,
0.9083541339 for dense-sparse and 0.9451486284 for the XLSX one-edit control.
The reopen phase is outside the workflow timer. Phase quantiles are
descriptive and do not sum to whole-operation quantiles.

The allocator lane is diagnostic. Publication allocation calls and bytes fell
as follows:

| case | calls baseline → candidate | allocated bytes baseline → candidate |
| --- | ---: | ---: |
| XLSX one-percent, medium | 19,248 → 416 | 3,970,346 → 2,765,098 |
| XLSX one-percent, dense-sparse | 36,366 → 416 | 5,575,801 → 3,275,001 |
| XLSX one-edit, medium | 2,662 → 308 | 1,216,622 → 1,065,966 |

Other phase allocation calls and bytes, and reallocation counts, are unchanged.
The measured live and peak fields are three bytes higher in the candidate at
the corresponding phase snapshots, so this lane does not establish a peak
live-byte reduction. The DOCX allocator control has no allocator metrics.

The [micro reduction](results/change-0905/micro-analysis.json) passes its hard
gate with no material adverse cases. The two cases with p99 review flags are
cardinality-2 authored and compact; their p50 ratios are 1.0371562973 with interval
1.0130410730–1.0521512106, and 1.0386657724 with interval
1.0200534324–1.0487967153. Both remain below the 1.05 p50 review threshold;
their p99 tail rows are retained as review flags. Six other jobs have smaller
p50 ratios above unity, retained in the full reduction. No micro spread flag fires.
The micro probe is a complete XML-audit batch and is separate from the Office
workflow benefit estimate.

The native process RSS medians are descriptively lower by roughly 6–7 MiB in
the measured cases. This is whole-process GNU `time` high-water data and is
affected by executable layout; it is not a causal heap or peak-live claim.
The [diagnostics](results/change-0905/diagnostics.json) record that the
allocator supervisor/preflight interval overlapped the final native supervisor
interval, while allocator workload children themselves remained serial. This
does not prove host exclusivity.

The host is a warm, serial,
fresh-process run pinned to CPU 12 on a 32-vCPU AMD EPYC 9R45 KVM guest under
Rust 1.95.0; other agents may share the host, so it is not a dedicated
machine. The full environment is retained in
[environment.json](results/change-0905/environment.json).

## Disposition and cleanup

The [retained replay](results/change-0905/replay.json) passes the frozen main,
micro, parity, source-archive and overlay checks. The independent result review
found no blocker, and [decision.json](results/change-0905/decision.json) accepts
the candidate with the documented tradeoffs. The packet retains the source
manifests, candidate patch, failed-attempt logs, parity JSONL, quality receipts,
all raw captures, reductions and replay evidence.

The eight local binaries under `results/change-0905/binaries/` (native,
allocator, parity and probe roles) are retained and gitignored. A fresh clone
does not contain them; replay custody instead binds their verified bytes to the
retained build receipts, while the source archive, four baseline overlays,
candidate audit overlay, manifests and command/toolchain receipts preserve the
rebuild path. [Cleanup](results/change-0905/cleanup.json) removed both marked
source/build roots and 7,379,644,211 logical bytes, unlinking fixture symlinks
without following them. [Post-cleanup replay](results/change-0905/post-cleanup-replay.json)
passes. No worktree was created for this batch. A separate branch commit is
prepared through an isolated Git index so the ongoing unrelated cherry-pick
and its staged/conflicted files remain untouched; target-branch integration
is still pending.
