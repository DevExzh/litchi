# 0916 — DOCX section commit target reuse

This batch removes redundant target parsing in the source-backed existing-
section layout edit/save workflow. It reuses the validated target snapshot
already retained by `layout::Commit` after normal current-source authorization.
The reviewed candidate passes correctness and the predeclared native gate;
pooled elapsed p50 falls from 4.840903 to 4.517335 ms on the named fixture.
The 0915 resolver-borrow candidate remains unadopted.

## Work and ownership

`Edit::commit` constructs the target main-document XML, parses and validates its
section inventory, checks topology, and verifies semantic readback. The commit
retains that immutable target. Its section inventory uses shared immutable
allocations, so cloning it does not reconstruct descriptors or XML scans.

Previously, commit publication delegated to patch publication, which obtains a
fresh current-source snapshot, authorizes the patch, and reparses the retained
target bytes. The adopted commit-only route keeps the current-source snapshot
and every authorization check, then reuses the commit target. Standalone patch
publication continues parsing its target because patches do not retain a parsed
snapshot. No-op commits must return the current snapshot, as `Patch::apply` does.

The [design review](results/change-0916/design-review.md) requires exact byte,
lineage, revision, artifact fingerprint, and policy checks in their existing
order. Full-artifact fingerprinting, source freshness fences, execution checks,
signature handling, the OPC preservation writer, output fingerprinting, and
inverse authorization remain mandatory. This introduces no parsed cache,
retained state, public API, dependency, or unsafe code.

The candidate retains both paragraph-cache trim calls from the old delegated
route and the original typed operation context. The shared writer receives
either the normally parsed patch target or the authorized commit target;
its preservation and freshness behavior is unchanged.

## Profile evidence

The [profile packet](results/change-0916/profile-README.md) isolates the timed
publication caller from setup and post-timer controls. In the full publication
capture, target `Snapshot::from_shared_xml` accounts for 6,391,431 of
240,572,558 guest instructions (2.656758%). A separately isolated `Patch::apply`
capture attributes 6,381,596 of 6,384,252 instructions to that target parse.
Both support the duplicate-work mechanism. Neither establishes its native
latency share: guest software hashing differs from native execution.

`python3 -B docs/performance/results/change-0916/profile-replay.py` verifies
the retained packet and its 36 referenced artifacts. Large raw captures remain
outside Git with file sizes, hashes, and capture commands in the reduced record.

## Native protocol

The [premeasurement protocol](results/change-0916/native-protocol.json) defines
eight alternating baseline/candidate process pairs, 30 warmups and 200 measured
samples per child on CPU 12. The baseline is the retained 0914 release binary;
controls are freshly measured, not taken from historical timings.

Admission requires the median paired baseline-p50/candidate-p50 ratio to reach
1.05, with a lower 95% bootstrap bound above one (10,000 process-pair draws,
seed 916). Pooled p50/p95/p99 and whole-process maximum RSS increases above 5%
are review flags. All 0914 corpus, source/output, preservation, refusal, phase,
cache, logical read, and sink gates remain required for each child.

Before any native capture, protocol revision 2 clarified that the coordinator can hold
this batch’s quality/build/profiler work idle, but cannot guarantee a quiescent
shared machine. The initial protocol bytes and their hash are retained. No native
measurements informed that correction.

Elapsed time sums six phase clocks and excludes inter-phase gaps, verification,
and teardown. Each iteration uses a fresh in-memory source within a warmed
process. Logical reads are not physical I/O; process maximum RSS includes setup
and verification and is not operation peak-live memory. Guest instruction
profiles do not establish native SHA cost or predict native speedup.

## Native result

All 16 scheduled children completed; all 3,200 measured samples are retained.
The median paired baseline-p50/candidate-p50 ratio is **1.074135880**, with a
95% bootstrap interval of **[1.059448719, 1.086657320]**. Every pair exceeds
1.05. The frozen timing gate passes. No pooled latency or process-RSS review
flag is raised. The [reduction](results/change-0916/reduced.json) retains every
child, pair and phase statistic; the [review](results/change-0916/review.md)
records independent verification.

| Metric | Baseline | Candidate |
| --- | ---: | ---: |
| Pooled elapsed p50 | 4.840903 ms | 4.517335 ms |
| Pooled elapsed p95 | 5.001463 ms | 4.607341 ms |
| Pooled elapsed p99 | 5.104194 ms | 4.717232 ms |
| Publication p50 | 3.017034 ms | 2.686882 ms |
| Query p50 | 1.337906 ms | 1.342611 ms |
| Commit p50 | 0.350226 ms | 0.354642 ms |
| Median whole-child maximum RSS | 37,646 KiB | 37,768 KiB |

Pooled p50 falls 6.684%, while whole-child RSS rises 0.324%. The RSS result is
not an allocation-count or peak-live-memory measurement. Query and commit
medians rise slightly; they are disclosed rather than hidden by the total.
Publication accounts for the observed latency reduction. This comparison uses
fresh controls and does not compare native timing against the software guest
instruction counts.

Every measured iteration still performs 202 logical source reads returning
8,409,488 bytes, retains one 18,855-byte main part before publication, and emits
2,101,446 bytes in 75 writes with a maximum write of 65,536 bytes. The output
hash remains `5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec`.
All semantic, untouched-member, inverse, cache, source and sink guards pass.
No sample was excluded or replaced.

The [capture barrier](results/change-0916/capture-barrier.json) was checked
after coordinated build, profiling and quality work ended and before the
manifest freeze. Child receipts retain before/after host load. Unrelated
shared-host work remains uncontrolled; this is not isolated-machine evidence.

## Scope

This is one generated DOCX edit/save fixture on one shared host. Producer
coverage, physical cold/range input, allocation-byte and peak-live measurements,
hardware counters, and parallel scaling remain outside the experiment. The
broader non-iWork performance goal remains active.

## Correctness scope

The isolated candidate's default-feature DOCX all-targets suite passes 1,914
tests, with no failures and one ignored test. The added unit checks changed
target XML/inventory allocation reuse and no-op sharing of the current
snapshot. Integration checks cover a changed unused payload with a stable
source version (fingerprint conflict before output) and a revision change
during capture (source-changed refusal before output). Existing tests and the
benchmark controls cover semantic readback, exact inverse restoration,
untouched ZIP members, signed input, and partial sink failure.

The authorization helper is an exact extraction of the existing checks, and
the common publication writer retains its existing behavior. The review does
not claim exhaustive new tests for every unchanged error branch. The static
boundary check passes for 65 packages and 244 internal dependency declarations,
with 11 explicitly recorded existing debt items. This is scoped DOCX and
boundary validation, not whole-workspace certification.

DOCX doctests also pass (78 passed, 31 ignored), as do formatting, Clippy with
warnings denied, and warning-denied rustdoc. The [quality receipt](results/change-0916/quality-receipt.json)
binds the exact commands and logs. The [evidence README](results/change-0916/README.md)
documents source/build custody, retained external artifacts, reproduction, and
replay. The full source manifest and compiled executable remain outside Git.

## Adoption

The independent review approves the exact retained three-file patch. The two
production files and focused tests are adopted byte-for-byte from the tested
candidate; no further optimization is bundled. The runner's immutable
`reduced.json` intentionally reports timing eligibility only. This change
record and the [adoption receipt](results/change-0916/adoption.json) record the
coordinator's decision after source, quality and evidence review.

The [cleanup receipt](results/change-0916/cleanup.json) records removal of the
owned extracted source and build caches while retaining the final executable,
source archives, raw evidence and receipts. Native and profile replay both
pass after cleanup.
