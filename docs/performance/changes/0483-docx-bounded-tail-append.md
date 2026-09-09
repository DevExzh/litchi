# 0483: bounded DOCX plain-paragraph tail append

The source-backed DOCX package can now append one caller-supplied plain
paragraph without retaining the complete main-document XML, candidate XML or
paragraph range vector. The operation scans through finite parser windows and
passes a bounded authored fragment and authenticated scalar proofs to OPC.
Publication regenerates the selected main member and preserves untouched raw
ZIP members. It inserts before an admitted final opaque section-properties
span, or immediately before the body close when no such span exists.

This is a new public route compared with the existing materialized paragraph
copy operation in the same executable. It is not a replacement for arbitrary
DOCX editing. Strict and Transitional sources, finite text and XML limits,
settings/MCE validation, package execution budgets, source freshness,
cancellation, exact no-ops and immediate exact inverse publication retain
explicit contracts. Unsupported main-story grammar and dependent topologies
fail closed. The [source checkpoint](../results/change-0483/docx-source-checkpoint.md)
and [ADR matrix](../results/change-0483/adr-matrix.md) describe those boundaries.

The [evidence bundle](../results/change-0483/README.md) uses deterministic
sources with 64, 8,192 and 131,072 existing paragraphs and one opaque 32 KiB
member. Both routes append the same text. Their generated XML differs in
namespace and whitespace declarations; independent semantic and raw-member
oracles authenticate each expected output. The corpus has no settings Part,
so these measurements do not establish a settings/MCE performance result.

The accepted matrix contains 24 fresh processes and 720 retained samples,
with separate normal and allocator binaries, three warmups, 30 samples per
process and CPU 2 affinity on the recorded AMD EPYC 9R45 machine. Git source
checkpoint `59b1b27ce` and compilation manifest `30d3ad60…` bind both binaries.
Normal total-lifecycle means are:

| Existing paragraphs | Materialized ms, R1 / R2 | Bounded ms, R1 / R2 | Bounded change, R1 / R2 |
| --- | ---: | ---: | ---: |
| 64 | 0.159472 / 0.157844 | 0.315602 / 0.314569 | +97.904% / +99.292% |
| 8,192 | 14.913749 / 14.921802 | 30.233109 / 30.469390 | +102.720% / +104.194% |
| 131,072 | 243.769324 / 240.863305 | 481.701301 / 485.276936 | +97.605% / +101.474% |

The three-size normal latency geometric ratio for R1 is 1.993959×.
The three-size normal latency geometric ratio for R2 is 2.016432×.
These ratios summarize only this one-paragraph route comparison. Every normal
latency pair regresses by about 98–104%; no normal latency repeat drift crosses
5%. The [complete tables](../results/change-0483/measurements.md) include
intervals, tails, all individual metrics and review flags.

Every bounded allocator sample has a 609,875-byte incremental operation peak,
292 allocation callbacks and 1,956,965 requested allocation bytes, with zero
net live-byte delta at exit. Materialized peaks are 506,786, 2,669,850 and
35,371,290 bytes across the three sizes. The bounded peak is therefore 20.34%
higher on the smallest source, 77.16% lower on the middle source and 98.28%
lower on the largest. Small-source requested allocation bytes also increase
89.60%. Allocator callback bytes include full realloc request sizes; they are
not physical copies or memory-bandwidth measurements.

Normal whole-process RSS remains about 100 MiB for both largest-source routes.
It includes corpus construction and independent readback, so the operation heap
reduction does not establish a whole-process RSS reduction or a general DOCX
memory bound. The caller-owned source, package metadata, codec state and
configured settings model remain distinct costs.

Logical source reads also expose the tradeoff. At 131,072 paragraphs, calls
increase from 62 to 119 and returned bytes from 1,033,651 to 2,056,921;
sequential sink write calls increase from 39 to 50. At 64 paragraphs, returned
bytes decrease from 8,077 to 5,773 while calls increase from 42 to 59. These are
adapter/sink observations, not filesystem or network syscall measurements.
Changed main XML is regenerated; untouched compressed members retain their raw
representation under the checked preservation contract.

The separate [process profiles](../results/change-0483/profile-review.md)
record 1.814× cycles and 1.934× instructions for bounded versus materialized,
while generic cache misses fall to 0.430×. These whole-process diagnostics
include setup, warmups and oracles; they are not operation PMU deltas.
Repeated XML auditing and replay are the measured CPU work to investigate next.
No speedup or parallel scaling claim follows from this batch.

The source checkpoint passed both DOCX feature configurations, warning-denied
Clippy and rustdoc, the non-iWork workspace and example checks, crate boundaries,
five focused harness tests, LibreOffice text readback and two 10,000-iteration
ASan/libFuzzer smoke runs. Historical failures are retained and classified.
The full workspace checks exposed iWork-only failures; the user explicitly
excluded that work from this performance program.

Evidence verification passes all 24 captures, 32 required gates, 12 pilots,
129 retained validation receipts and the 62-seed tail fuzz chain. Independent
recalculation matches all summary statistics and flags. After runtime binary
cleanup, copied-bundle verification and seal replay pass, and all eight
corruption cases are refused. The [closure record](../results/change-0483/evidence-validation/closure-check.json)
retains the command outputs. The first formal attempt and initial bytecode
cleanup refusal remain archived; neither is presented as accepted evidence.

The change introduces no parallel execution, ambient source provider or
filesystem behavior. A finite single paragraph does not satisfy the goal's
very-large authored-stream requirement. Replayable paragraph events, durable
forward patches, explicit original-source inverse recipes, broader source/sink
and cold/warm/concurrent measurements remain required. The complete non-iWork
CRUD objective stays open; [next work](../results/change-0483/next-work.md)
preserves its scope.
