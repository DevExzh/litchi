# 0497 evidence bundle: atomic DOCX publication

## Status

The candidate, plan, and review inputs define the atomic filesystem route and
the measurement contract. Formal1 capture and verification are complete: all
288 child terminals report `status: pass` and `exit_code: 0`, and the formal
analysis retains 8,640 samples. The resumed suffix completed 97/97 children
with exit code 0. Final cleanup verification passes. `seal.py` records the
complete evidence inventory and verifies it before commit.

The full non-iWork performance goal remains open. The change is scoped to the
bounded DOCX logical-tail append route and adds consuming filesystem methods for
atomic publication. The default hashing route remains the only before/after
comparison. Counting and atomic routes are after-only capability
records, and no general speedup or atomic performance claim follows.

## Evidence and review inputs

- [plan](plan.md) defines the API, timing boundary, planned matrix, validation,
  and retention contract.
- [production review](production-review.md) records the source ownership,
  destination alias, source-fence, failure, and `Committed` semantics.
- [harness review](harness-review.md) records the route-schema, failed-child
  custody, syscall path/order, and destination-inverse boundaries.
- [measurement review](measurement-review.md) records the route separation,
  ABBA ordering, report schema, and descriptive analysis boundary.
- [section-boundary review](sections-failure-review.md) explains the existing
  tight-budget refusal and the test-only finite-budget event-limit overlay.
- [validation plan](validation-plan.md) and [profile plan](profile-plan.md)
  define the remaining gates and descriptive profile limits.

The current candidate source manifest and fixture verification are retained in
`candidate-source.json` and `fixture-verification.json`. The frozen protocol,
four normal/allocator build receipts, and pilot1 verification are retained as
pre-formal evidence. The original ordinal-191 observation is archived under
`interrupted/formal1-enospc` without a fabricated report, terminal, or resource
receipt. The reviewed resume reran ordinal 191 through 287 in the same frozen
order and completed all 97 children. The formal child inventory and terminal
custody are verified by [verification/formal1.json](verification/formal1.json);
the descriptive result is retained in
[analysis/formal1.json](analysis/formal1.json).

## Current validation

The retained final-gates receipt records 16 named gates passed and four
normal/allocator binaries built. Its current test counts are 1,363 DOCX
default tests, 1,382 DOCX feature tests, 33 focused tests, 627 OPC preservation
tests plus one ignored test, six OPC atomic tests, 79 doctests plus 31 ignored,
and 483 harness allocator tests plus one ignored. Pilot1 verification passes
72 children and 216 samples. ASan runs with seeds 497 and 498 each pass 10,000
cases. Formal1 capture, terminal custody, analysis, and formal verification
pass within the documented descriptive scope.

The immutable strace1 attempt failed before benchmark launch because the host
rejects the requested `fstatat` syscall; strace2, strace3, and strace4 retain
typed or host-filter failures. The `profile_v2.py` strace5 attempt passes and
is verified as diagnostic evidence for the 8,192-source/64-authored near-text
window/file-store/data case with one warmup and one sample. It makes no
operation-attribution or performance claim. Early target cleanup removed
12.496 GiB and completed the Cargo target cleanup; the fuzz-target cleanup
removed 1.070 GiB; separate external cleanup receipts removed 1.865 GiB and
27.117 GiB. Final cleanup verification passes after removing 512,319,488
allocated bytes and 20,515 files; only five replay executables remain.
`cleanup-attempt1.json` preserves the initial partial cleanup, and
`cleanup.json` records recovery after removing the empty profiling parent.
The process audit records inaccessible foreign processes as a visibility
limitation and does not claim host-wide quiescence.

The formal run records an emergency shared-disk cleanup overlap. Whole-child
intervals for
`r2-after-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`
and
`r2-before-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`
overlapped that cleanup interval, which removed 16,923,615,232 bytes and
created potential I/O interference. The bundle is not treated as isolated-
hardware evidence, and individual outliers are not attributed to that cleanup.

The pre-existing section event-limit assertion is preserved as a typed 16 MiB
memory-admission refusal plus a test-only finite 256 MiB event-limit case. The
production scanner limits and reservations are unchanged; the failed
development receipt remains retained for audit.

## Captured shape and analysis boundary

The frozen inventory names 18 explicit provider/workload arms in normal and
allocator roles. Two reversed repeats use three warmups and 30 measured samples
per child. The formal matrix contains 72 before/default hashing children, 72
after hashing children, 72 after counting-sink children, and 72 after
atomic-path children: 288 children and 8,640 samples. Pilot1 verification
passes for 72 children and 216 samples with one warmup and three samples; it is
excluded from formal analysis. The formal verification receipt reports all 288
children pass, including the 97-child resume. The original interrupted
ordinal-191 observation remains an archived unknown custody record and is not
reconstructed.

The analysis retains 864 matched default-hashing comparison cells and 144
after-only capability rows. It retains all 100 cells above the five-percent
adverse threshold: 72 allocator live-byte endpoint, 16 latency, and 12
whole-child RSS flags.
Eighty-three are aggregate two-repeat flags and 17 are single-repeat flags;
none are dismissed or converted into a causal claim. These are descriptive
capture counts, not a broad speedup result.

The 72 allocator live-byte flags are endpoint observations, not allocation
count or operation-delta results. A possible harness-level explanation is
that the current sample representation reserves an inline
`Option<PublicationRecord>` before the measured region even when it is
`None`; this is an inference about harness layout and makes no production
live-memory claim. The file-store small allocator hashing p99 remains
unresolved at +73.58% and +75.90% across the two repeats (aggregate median
+74.79%), with no causal attribution.

The default hashing route retains its existing report shape and is paired
before/after within matching arm and role. The counting route is a bounded,
non-retaining after-only sink with production artifact proof. The atomic route
is an after-only filesystem capability; its timing includes sibling temporary
creation, publication, file/data synchronization, replacement, parent-directory
synchronization, publication drop, and the existing lifecycle. Readback,
semantic/raw checks, fixture inverse comparison, replay cleanup, report writing,
and process/allocation endpoint snapshots are outside that interval.

Atomic write-call and sink-digest values are unobservable in this design and
must remain absent, empty, or null. The production publication fingerprint and
length authenticate the counting and atomic outputs. Any future analysis must
keep route capability observations separate from default hashing comparisons.

## Acceptance boundary

The implementation reviews establish the candidate contract, and formal1 now
has retained child terminal/RSS/allocation/read/replay/output/cleanup evidence,
correctness and conservation oracle validation, regenerated analysis, and
formal verification. Final cleanup verification passes; `seal.py` inventories
and revalidates the evidence. No cross-platform Windows guarantee, crash-durability result, native
Word producer result, cold filesystem result, or broad CRUD result is implied
by this bundle.
