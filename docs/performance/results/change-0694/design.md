# Expanded MCE element-name ownership

performance_claim: none

## Hypothesis and scope

Baseline: `829bed696`. After 0693 removed redundant per-slide processing,
the fresh native open-plus-capture profile still places 12.48% of self samples
in shared `mce::codec::start`. Its eager `expand` allocates a namespace String
and local-name String on every start element, including the default empty
extension profile. Most decisions only compare namespace/local slices.

The candidate calls the same `expand_parts` at the same point, after source
attribute decoding and local namespace installation and before directive
validation. It retains borrowed namespace/local names. An owned Name remains
necessary for the existing hashed extension lookup when the extension set is
nonempty, and for preservation/process matching on an ignorable element.
These cases keep the existing algorithms, with at most one lazy owned Name
per element. Inherited opaque descendants keep QName validation but avoid
ownership that their early-return branch never consumes.

## Constraints

| Accepted decision | Application |
| --- | --- |
| 0001 / 0004 | No public API or error vocabulary changes; ownership remains private. |
| 0002 / 0010 / 0011 / 0024 | Shared MCE remains in its existing common owner; no new dependency. |
| 0003 | No snapshot, transaction, patch, conflict, or publication changes. |
| 0005 / 0030 / 0031 | No cache, I/O, concurrency, execution, or resource policy change. Temporary allocation work is measured separately from native time. |
| 0006 | Identical QName checks, processing decisions, reports, output bytes, and typed error precedence required. |
| 0008 | Owner and consumer checks, tests, lint, docs, boundary, and evidence gates must pass. |

All 33 previously read GOAL/ADR constraint hashes match the preceding batch;
the baseline manifest freezes them again. There is no unsafe production code,
SIMD, change to container bytes, or ambient service. Removing infallible String
allocations removes those particular process-abort opportunities; this is not
a claim of recoverable allocator exhaustion or an exact RSS bound. Existing
finite parser limits and fallible reservations stay at their original sites.

## Evidence plan

The native and allocator executables are separate. The unchanged phase probe
measures 13 PPTX one/no-op/two-edit cases on real, generated, marker-control and
notes-bearing sources. Source materialization and target derivation are outside
the timers; full save latency is not measured. Two baseline A/A legs precede
A/B/B/A with 100 samples and 5 warmups per case/leg, pinned to CPU 12 on a
shared warm host. Retain every phase, tail, confidence interval and >5% trigger.
The marker-stripped archive is only a mechanism control, not a semantic oracle.

The ten-case public refusal probe checks exact graph/error metadata and timed
capture refusals. The independent common-MCE oracle compares output digest,
length, ownership and Report, or exact typed Debug errors, across
real OOXML parts, deterministic mutations, explicit opaque extension profiles
and limits. Owner regression tests and DOCX/XLSX/PPTX consumer suites complement
this bounded oracle; no universal fuzzing or cross-platform claim follows.

Retain allocation calls/requested bytes/peak live changes; use native perf
sampling and counter slopes to assess instructions and cycles. Price native
text growth and process RSS separately. No native timing is taken from the
counting-allocator executable.
