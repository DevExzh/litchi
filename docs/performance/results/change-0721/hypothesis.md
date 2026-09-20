# 0721 — bounded DOCX structural fusion pilot

The 0720 trace binds the current `DocumentBody::from_xml` boundary: the
ordinary-save generated-medium body contains 21,517 bytes / 1,410 events per
structural walk; NumberedList contains 4,563 bytes / 178 events. Both pay two
complete structural reader traversals before body capture. Their independent
MCE inputs are respectively 0/200 and 0/5 offsets. These counts are a work
hypothesis, not a native saving forecast.

The candidate shares only the reader traversal. Alt metadata and block-range
state retain their separate namespace predicates, depth/node/anchor counters,
capture rules and errors. An observer failure is deferred through completion
of the alt parser and first MCE selection. Both MCE calls retain their exact
individual inputs and order. The final body reader and source-backed consumer
remain unchanged. Owned events and resolver clones remain part of the alt
parser; this does not revive rejected borrowing or section-collection pilots.

Correctness gates include independent legacy range-state differential tests,
public facade/error/publication parity, actual per-call MCE trace equality,
source bounds, exact output identities and DOCX tests/Clippy/docs. The tracing
lane must count reader calls separately from state-machine observer calls so
one shared event cannot be mislabeled as two reads.

Native edit and lifecycle samples use existing ordinary-save selectors. Edit
covers `owner.edit` with open outside; lifecycle covers open, edit and atomic
save. Owner destruction, output readback and cleanup are outside both clocks
and allocation regions. These phase times must not be added. Allocator net-live
and peaks are per-operation diagnostics, not leak or process-RSS measurements.
The source snapshot contains the exact current timer implementation.

Freeze two ABBA cycles, fresh child per corpus/phase/stage, 100 warmups and 200
native samples. Every paired edit p50/mean must improve at least 3%; every
lifecycle p50/mean must regress no more than 3%. The first cycle separately
captures allocator counters at 0 warmups/3 samples, with request count and
requested bytes bounded to 3% regression. Retain all >5% tail/repeat flags;
never discard or resample a slow child to obtain acceptance. Read-only controls
qualify the shared range helper's other consumers under their separately frozen
plan. No hardware, RSS, cold-cache, throughput or scaling claim is made.

Interleaving bounded range storage changes allocation lifetimes. Preserve
fallible-reservation resource names, explicit limits and error precedence;
this pilot cannot prove identical host allocator-exhaustion scheduling. Reject
any deterministic resource-limit or preservation regression. Revert production
if any correctness or frozen performance gate fails; retain reproducible
negative evidence and useful independent tests where applicable.
