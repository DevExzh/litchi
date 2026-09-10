# 0497: atomic DOCX logical-tail publication

0497 adds a format-owned filesystem publication capability to the existing
bounded, replayable DOCX logical-tail append route. The current candidate and
architecture reviews establish the API and custody boundary. Formal1 now
retains 288 status-pass child terminals and 8,640 measured samples: the
original 191-child prefix was resumed at ordinal 191, and the 97-child resume
completed with exit code 0. The formal analysis and verification are complete
within their descriptive scope, and final cleanup verification passes.

The full non-iWork performance goal remains open. This change supplies an
atomic-save enabler for one bounded DOCX append route. It does not establish a
general atomic-save result, a speedup, a native-producer round trip, genuine
borrowed-source lifetime, cold intersection, parallel scaling, arbitrary
logical-tail append, durable history/composition, or broad CRUD/security
coverage.

## Current implementation

`ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` are consuming methods over the existing
plan and commit products. They delegate the existing source-backed splice to
OPC's atomic sibling-temporary replacement helper. The callback writes the
candidate to a private sibling, performs the existing source, candidate,
authored-replay, budget, and cooperative-cancellation checks, and only then
allows the filesystem layer to synchronize the temporary, replace the
destination, and synchronize the parent directory.

Before replacement, source, replay, limit, cancellation, sink, and late
managed-output failures preserve the prior destination and remove the private
temporary. A successful replacement followed by a parent-directory sync
failure returns `OpcError::Committed`; the consuming API returns no publication
or inverse product in that state, so callers must inspect the destination and
must not blindly retry. Source freshness is a final cooperative callback fence,
not a compare-and-swap or continuous source watch.

The focused route cases cover new and existing destinations, exact bytes and
reopening, durable forward/inverse authorization, replacing the open source
path, and a Unix hardlink alias. They also cover symlink and nonregular
destination refusal, source mutation, cancellation, a replay failure after
preparation, late output-budget failure, permission behavior, temporary cleanup,
and managed workspace release. Unix alias behavior does not establish a
Windows same-path or hardlink guarantee. The existing OPC unit test remains the
injected evidence for the post-replacement parent-sync `Committed` branch.

The candidate also carries a test-only section-boundary overlay. The unchanged
16 MiB fixture continues to prove typed memory admission refusal; a separate
finite 256 MiB test budget admits the deliberately large event fixture so the
independent event-limit refusal is reached. Production scanner limits and
reservations are unchanged.

## Publication and measurement boundary

The historical default hashing sequential sink remains the default route and
keeps its existing report shape and before/after timing scope. The new
counting, non-retaining sequential sink is an after-only capability baseline;
it uses the production publication artifact length and fingerprint rather than
inventing a sink digest or write-call count. The atomic destination is also an
after-only capability baseline because the before revision has no atomic
destination method. Its timed interval includes source admission and the
production sibling-file write, data synchronization, replacement, parent
directory synchronization, publication drop, and the existing lifecycle work.

Atomic destination readback, semantic and raw-member checks, fixture inverse
comparison, replay cleanup, report serialization, and process/allocation
endpoint snapshots remain outside the elapsed interval. Atomic reports expose
no synthetic write-call or digest counters; optional sink fields remain empty
or null. The artifact proof comes from the production publication result, and
the atomic route's inverse scope is the untimed fixture-publication inverse
oracle unless a later capture adds a destination readback inverse replay.

The frozen protocol has 18 explicit provider/workload arms in normal and
allocator roles, two reversed repeats, three warmups, and 30 measured samples
per child. It defines 72 before/default hashing children, 72 after hashing
children, 72 after counting-sink children, and 72 after atomic-path children:
288 children and 8,640 samples. Pilot1 verification passes for 72 children and
216 samples and is excluded from formal analysis. The original formal1
coordinator stopped at ordinal 191 after 191 completed children because it
encountered `OSError: [Errno 28] No space left on device` while writing
replay-cleanup custody. That raw ordinal-191 observation remains archived
without a fabricated report, terminal, or resource receipt. The reviewed
resume reran ordinal 191 through 287 in the same frozen order; its
`resume-terminal.json` reports 97/97 completed children with exit code 0.
Every formal child now has a passing terminal, report, resource, and replay
cleanup receipt. `verification/formal1.json` passes, and `analysis/formal1.json`
contains the 288-child, 8,640-sample descriptive analysis.

The analysis retains 864 matched default-hashing comparison cells and 144
after-only capability rows (72 counting-sink and 72 atomic-path). It keeps all
100 cells above the five-percent adverse threshold: 72 allocator live-byte
endpoint, 16 latency, and 12 whole-child RSS flags. Eighty-three are aggregate
two-repeat flags and 17 are single-repeat flags that do not aggregate above the
threshold; none are dismissed or resolved as causal evidence. These figures
describe this frozen capture and do not authorize a broad speedup claim.

For one normal deterministic-owned slice, the descriptive p50 latency figures
are below (milliseconds, repeat one / repeat two):

| Workload | Hashing before | Hashing after | Counting after | Atomic after |
| --- | ---: | ---: | ---: | ---: |
| `s64-a64` | 0.553 / 0.547 | 0.565 / 0.566 | 0.564 / 0.565 | 5.674 / 5.510 |
| `s64-a16384` | 77.407 / 77.488 | 79.747 / 78.936 | 78.639 / 79.409 | 91.149 / 84.652 |
| `s131072-a64` | 309.207 / 307.985 | 311.234 / 308.519 | 308.348 / 309.652 | 319.784 / 318.130 |

These rows are descriptive full-lifecycle route observations for the named
slice. Hashing before/after is the only matched comparison; counting and
atomic are after-only capabilities with different sink/publication contracts,
so the table does not establish route ranking or a speedup claim.

The 72 allocator live-byte flags are endpoint observations, not allocation
count or operation-delta results. One possible harness-level explanation is
that the current sample representation reserves an inline
`Option<PublicationRecord>` before the measured region even when it is
`None`; this is an inference about harness layout and makes no production
live-memory claim. The file-store small allocator hashing p99 remains
unresolved: the retained repeat-one and repeat-two deltas are +73.58% and
+75.90% (aggregate median +74.79%), with no causal attribution.

The legacy before/after comparison is restricted to matching default hashing
rows: repeat one runs before then after, and repeat two reverses that order.
Counting and atomic routes are balanced against each other only within the
after-only capability block; they are never treated as an atomic before/after
comparison. Never fill unobservable atomic write-call counts with synthetic
measurements.

## Current review and custody state

The read-only production review records the candidate source, harness
identities, and ownership boundary. Its earlier test counts are retained as
historical review inputs; the current gate, capture, profile, and cleanup state
is recorded below.

The measurement review records 15 passing focused Python method tests and
checks the frozen ABBA inventory and route separation. The harness review
records atomic report-schema fields, failed-child artifact custody, exact
syscall path binding and ordering, and the distinction between fixture inverse
proof and destination inverse replay. The section-boundary review confirms the
pre-existing tight-budget failure in both before and candidate trees.

The retained final-gates receipt records 16 named gates passed and four
normal/allocator binaries built. Its current test counts are 1,363 DOCX
default tests, 1,382 DOCX feature tests, 33 focused tests, 627 OPC preservation
tests plus one ignored test, six OPC atomic tests, 79 doctests plus 31 ignored,
and 483 harness allocator tests plus one ignored. Pilot1 verification passes
72 children and 216 samples; ASan runs with seeds 497 and 498 each pass 10,000
cases. Formal1 capture, terminal custody, analysis, and formal verification
now pass within the scope above. The immutable strace1 attempt failed before
benchmark launch because the host rejects the requested `fstatat` syscall;
strace2, strace3, and strace4 retain typed or host-filter failures. The
`profile_v2.py` strace5 attempt passes and is verified as diagnostic evidence
for the 8,192-source/64-authored near-text window/file-store/data case with one
warmup and one sample; it makes no operation-attribution or performance claim.
The tight 16 MiB typed refusal and test-only finite 256 MiB event-limit case
remain unchanged in production.

The formal run also records an emergency shared-disk cleanup overlap. Whole
child intervals for
`r2-after-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`
and
`r2-before-hashing_sink-allocator-deterministic-latency-s64-a16384-short-c64`
overlapped the cleanup interval; that receipt records 16,923,615,232 bytes
removed during the interval. This creates potential I/O interference, so the
run is not treated as isolated-hardware evidence and individual outliers are
not attributed to that cleanup.

Early target cleanup removed 12.496 GiB and completed the Cargo target cleanup;
the fuzz-target cleanup removed 1.070 GiB. Separate external cleanup receipts
removed 1.865 GiB and 27.117 GiB. The earlier external cleanup preceded formal1;
the overlap above is recorded separately. Final cleanup removed 512,319,488
allocated bytes (20,515 files) and retains only the five replay executables.
The initial cleanup stopped after deleting its candidates because an empty
profiling parent remained; `cleanup-attempt1.json` preserves that failure,
and `cleanup.json` records the explicit recovery and verified final state.
The process audit checks same-user and accessible processes; inaccessible
foreign processes are recorded as a visibility limitation, with no host-wide
quiescence claim. Arbitrary foreign command lines are not retained.

The source and fixture manifests, build records, reviews, analysis, formal
verification, and validation plans are retained under
[`results/change-0497`](../results/change-0497/). The protected
`/home/zhuhe/code/litchi-spec-gaps` worktree is excluded. Cleanup verification
passes. `seal.py` inventories and revalidates the retained evidence; its
custody scope does not change the measurement scope.

## Remaining boundaries

The atomic route is a necessary filesystem-publication capability for this
bounded logical-tail operation. It is not a durability or cross-platform
guarantee beyond the tested contract. The formal figures are descriptive
full-lifecycle observations; the default hashing route is the only
before/after comparison, counting and atomic routes remain after-only, and no
causal or broad speedup claim is authorized. Borrowed sources, independent
producers, cold behavior, bounded concurrent execution, arbitrary append
semantics, durable history/composition, and the wider non-iWork CRUD/security
matrix remain open.
