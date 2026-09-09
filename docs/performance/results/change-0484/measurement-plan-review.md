# Change 0484 measurement-plan review

Status: **validator revision reviewed; formal bundle still unsealed**. This is
a read-only acceptance review of the current measurement driver and replayable
tail-append harness. I did not change Rust, the driver, the protocol, or any
source file, and I made no timing claim. The review covers the
source/authored/chunk axes, the timed lifetime, counter and oracle custody, and
the acceptance gaps that must be closed before formal captures can support a
measurement result.

The retained `stream-harness-tests-dev46` receipt reports all 11 focused
harness tests passing with unchanged source. The `stream-cli-smoke-dev50`
receipt is useful shape/oracle evidence for the 64-source/64-authored,
fixed-64-byte, short and near text cases, but it is a debug `cargo run` with
one warmup and one sample and exported fixtures. Its report is diagnostic only;
it is not a formal release capture and supplies no timing conclusion.

The former report-shell gap was corrected in the frozen validator revision
`stream_measure_validation_0484`. Its seven focused Python tests pass. The
follow-up below separates the checks that are now enforced from the smaller
remaining validation limits; the one-shot, input, sink-size, and compression
arm findings later in this document remain open.

## What the frozen deterministic arm would measure

The driver deliberately deduplicates a one-axis-at-a-time matrix to ten cases
(`measure.py:87-142`):

| Axis | Cases in the deterministic arm | What is held fixed |
| --- | --- | --- |
| Source paragraphs | 64, 8,192, 131,072 | 64 authored, short text, fixed-64 chunks |
| Authored paragraphs | 64, 256, 4,096, 16,384 | 64 source, short text, fixed-64 chunks |
| Chunk partition | one, 64-byte, replay-window | 64 source and authored, near-limit text |
| Text mode | empty, short, near-limit | 64 source and authored, fixed-64 chunks |

The duplicate `(64, 64, short, 64)` and `(64, 64, near, 64)` points are
removed, leaving ten cases. The formal plan has two roles (`normal` and
`allocator`), two repeats, 30 measured samples and three warmups per process;
that is 40 fresh child processes and 1,200 measured samples
(`measure.py:171-234`). Repeat two reverses the case order, which is useful for
drift detection. The protocol explicitly authorizes no before/after speedup
claim (`measure.py:196-205`).

This is attributable coverage, but it is not a full factorial matrix. There is
no formal case combining a large source with a large authored stream, and the
near-limit mode is present only at 64 source and 64 authored paragraphs. The
results therefore cannot establish source-by-authored interaction, large
source plus near-limit text, or 16,384 authored near-limit behavior. Those are
limits of the claim, rather than reasons to relabel the ten cases as a full
cross-product.

The near-limit contract is concrete. `MAX_CURSOR_TEXT_BYTES` is 60 KiB and
`fill_text` fills exactly that bounded buffer for every near-limit paragraph
(`docx_replayable_tail_append.rs:44-57`, `:600-679`). XML escaping and the
paragraph envelope make encoded XML larger than 60 KiB; the report must retain
both raw authored text bytes per paragraph and encoded-byte/proof fields. A
near-limit encoded size must not be mistaken for the caller text limit.
The frozen `_check_authored_identity` now requires `near_limit` to equal
`authored_count * 60 KiB`, in addition to the empty/nonempty distinction and
chunk/event arithmetic (`measure.py:640-672`). The focused regression mutates
that payload and is rejected (`test_measure.py:281-302`), so the exact raw
near-limit-size acceptance finding is closed. The authored SHA fields remain
proof identities supplied by the child; the Python gate validates their format
and cross-field equality, rather than rehashing raw XML bytes that are not in
the report.

## Timed lifetime and what remains outside it

The harness states that corpus construction, XML/semantic/ZIP oracles, and
inverse checks are outside the timed lifecycle; the timed sample owns source
admission, preparation, publication, the short sink, and drops
(`docx_replayable_tail_append.rs:1-8`). The implementation matches that
statement. `Instant::now()` starts immediately before the inner operation
scope, and elapsed time is read after the scope has dropped the source-backed
package, edit, plan, generated provider, and sink
(`docx_replayable_tail_append.rs:1795-1834`). The scope includes
`Package::from_read_at`, `prepare`, `write_to_stream`, `HashingSink::finish`,
and the operation-owned drops. Fixture construction and all candidate/source
oracles happen once before sample loops (`:1714-1787`, `:1883-1934`).

Allocator-region acquisition and the before-process snapshot precede the
`Instant`; allocator-region finalization and the after-process snapshot follow
the elapsed read. That keeps instrumentation setup/teardown out of elapsed
time, but it also means the process and allocator records must not be described
as identical elapsed-time brackets.

The process record is also broader than the timed operation. The procfs delta
is a saturating before/after process counter difference, and its `rss_bytes` is
the saturating RSS delta while `peak_rss_bytes` is the absolute after-sample
VmHWM endpoint (`process_metrics.rs:59-87`, `:115-143`). The `/proc` probes
themselves can contribute I/O counters. The external GNU `time -v` resource
file is a whole-child-process maximum, including fixture/report setup and
serialization. It must be labeled as whole-process RSS and parsed separately
from operation RSS; it is not an operation-only resident peak.

## Counters, allocator semantics, and per-sample acceptance

The generated provider is intentionally borrowed and bounded: each cursor
uses one 60 KiB array and `open` increments a counter. The runtime check
requires exactly five opens and checks event and text-byte totals against the
independent authored proof (`docx_replayable_tail_append.rs:513-598`,
`:1836-1862`). The frozen capture validator now repeats those checks on the
serialized sample vector rather than relying only on the child-side assertion.

The former hard pre-freeze blocker is now closed for the serialized report.
`_check_report_shell` requires the sample list to have the requested cardinality
and calls `_check_sample` for every indexed sample
(`measure.py:754-837`). The seven focused tests cover valid cardinality and
rejection of truncation, forged proof/limit identity, allocator imbalance,
histogram fabrication, near-limit underfill, workspace mismatch, and non-finite
values (`test_measure.py:217-314`). For every sample in all 40
formal reports, the frozen validator now:

* require the measured sample count and indexes, positive elapsed nanoseconds,
  and exact source/authored/chunk/text identity;
* require source `calls` to be nonzero, `returned_bytes <= requested_bytes`,
  and each request/returned histogram to sum to `calls`;
* require authored `opens == 5`, event and text-byte totals equal the proof
  multiplied by five, and text-chunk totals agree with the selected chunk
  mode (the current child check does not validate `text_chunks`);
* require sink accepted bytes and SHA-256 to equal the candidate oracle,
  sink write-call histograms to sum to `write_calls`, `largest_write` to obey
  the configured 4 KiB bound, and all sink digests to be valid;
* require normal samples to carry the intentional uninstrumented allocator
  state and allocator samples to be complete, `Measured`, operation-scoped,
  non-overflowing observations; and
* verify allocator absolute invariants before deriving any increment:
  `peak_live_bytes_before <= peak_live_bytes_after` and
  `region_peak_live_bytes <= peak_live_bytes_after`, with the region peak also
  at least both operation live endpoints as required by
  `allocation_metrics.rs:215-265`.

Allocator byte totals are request accounting. A realloc increments allocation
and reallocation calls, adds the complete new-size request, and records the
old-size deallocation (`allocation_metrics.rs:547-565`). They are not physical
copy bandwidth. `region_peak_live_bytes`, `live_bytes_*`, and
`peak_live_bytes_*` are absolute process values; a region increment must be
computed as `region_peak_live_bytes - live_bytes_before` and never reported as
the raw absolute peak.

The latest validator closes the remaining counter checks: it rejects any
nonzero `failed_allocation_calls`, applies weighted histogram lower/upper
bounds, and requires `largest_write` to match the highest occupied bucket
(`measure.py:510-560`, `:762-786`, `:818-864`). The focused allocator and
histogram regressions cover those rejection paths (`test_measure.py:252-279`).
One operational evidence limitation remains: the process observation is
optional, so when the child cannot read procfs, `process: null` passes
`_check_process_sample` (`measure.py:789-801`). That is acceptable only if
formal reports treat procfs counters as best-effort and keep GNU `time`
resource evidence separate; it cannot support a claim that every sample has
in-process RSS/CPU counters.

## Source, authored, and artifact oracles

The source adapter owns an `Arc<[u8]>`, performs positional reads, and records
calls, requested/returned bytes, and both histograms
(`docx_replayable_tail_append.rs:352-426`). The generated authored proof
independently records event framing, escaped XML bytes, entity references,
chunk bound, replay window, and event/encoded SHA-256
(`:709-804`). The frozen validator now checks the source totals and histogram
conservation, five authored opens, event/text-chunk/text-byte multiples, sink
output identity, weighted bucket bounds, largest-write bucket consistency, and
the proof/oracle relationships (`measure.py:804-864`, `:704-759`). This closes
the prior “keys exist but sample values are unchecked” acceptance gap.

The fixture construction performs independent candidate XML and semantic
checks, source immutability, opaque-member equality, physical member order,
untouched raw local and central records, and exact immediate inverse
(`docx_replayable_tail_append.rs:1628-1711`). For untouched members the raw
local record, payload, descriptor, and every central byte are compared; only
the central-directory local-header relocation offset is normalized
(`:1168-1248`). This is the right preservation oracle. It is intentionally
outside the timed operation and is run once per child, not once per measured
sample.

The source archive is generated deterministically in the harness and its
source/candidate hashes are retained in each report. The frozen shell now
authenticates the internal source, authored, proof, oracle, event-bound, and
limit relationships for one report. Formal acceptance should additionally
require identical source, authored proof, limits, and candidate oracle hashes
for the same case across normal/allocator roles and repeats; the current shell
does not perform that cross-report equality check.

The per-report proof check still bounds `candidate_event_count` and requires it
to exceed `source_event_count`, but does not independently derive the exact
source/candidate event counts from the XML proof (`measure.py:646-653`). It
also validates hash formats and equality between proof and metadata rather than
rehashing archive bytes, which are not present in the report. Those are
appropriate reasons to retain the child-produced raw archive/oracle artifacts
and to run a cross-report/retained-artifact verifier before sealing a bundle;
they should not be described as cryptographic re-authentication of the raw
archives by the Python shell.

## Custody and binding

The custody path is structurally sound when all retained receipts remain in the
sealed bundle:

* `common.snapshot` records a content-addressed manifest of Rust, TOML, lock,
  and relevant XML inputs; `gate.py` captures it before and after each build or
  validation command (`common.py:80-135`, `gate.py:40-77`).
* `measure.py build` requires the gate's source identities to be equal, copies
  the release binary to an attempt-specific path, and records matching size and
  SHA-256 metadata (`measure.py:332-377`). `_load_builds` rechecks the copied
  binary and requires normal/allocator source identities to match
  (`:381-399`).
* Each capture binds the frozen protocol hash, build-receipt hash, copied
  binary metadata, exact argv, working directory, and controlled environment;
  output/resource/stdout/stderr hashes are retained in a non-replacing receipt
  (`measure.py:529-620`).

The custody path is not, by itself, sample acceptance. The protocol binds
helper hashes and the build receipt binds the source manifest; the protocol
does not contain a post-build source snapshot, so the retained build
`source_before`/`source_after` manifest and copied binary must travel with the
bundle. The resource file is retained but is not parsed by `_check_report_shell`.
The formal verifier must parse the GNU resource record, distinguish its
whole-process RSS from the in-process fields, and require every expected
receipt exactly once. At review time the change directory had no frozen
`protocol.json`; no build or capture may be accepted until `freeze` has written
it and all subsequent build/capture receipts bind its hash.

## Runtime scaling and ceiling interpretation

The limits are calculated from the generated source and authored proof, not
from a single fixed “workspace” constant. `stream_limits` sets source XML plus
2 MiB, candidate XML plus 2 MiB, output XML plus 8 MiB, paragraph and event
ceilings, authored text/XML/fragment ceilings, a 64 KiB patch ceiling, and a
replay window of at least 64 KiB (`docx_replayable_tail_append.rs:1033-1121`).
The authored event ceiling includes an entity-reference term, so XML-sensitive
near-limit text increases parser events even when raw text bytes per paragraph
remain fixed.

`scanner_workspace_limit` derives workspace from the maximum encoded token and
depth, including namespace, attribute, hash, and scope structures
(`:906-1031`). The report retains parser event limit, token bytes, workspace
bytes, depth, authored chunk bound, and replay window. The frozen validator now
recomputes the event and token limits exactly, mirrors the namespace/
attribute/scope workspace formula for the pinned 64-bit target, binds
depth/chunk/replay values to the authored proof, and enforces selected upper
bounds (`measure.py:562-601`, `:675-701`). The focused `+1` workspace
regression is rejected (`test_measure.py:304-308`), so the parser
workspace/event-ceiling acceptance blocker is closed for the pinned target.

The report still does not serialize every source/candidate/output/text/
paragraph ceiling from the Rust `stream_limits` object. Those omitted ceilings
remain a reporting limitation: they must not be presented as independently
measured dimensions unless a retained-fixture verifier derives them. The
selected parser workspace formula itself is now independently checked.

## Scope blockers against the 0484 design claim

The design's measurement plan requires both deterministic replay and an
explicit one-shot store arm, with memory storage distinguished from a bounded
caller-owned/external window (`design.md:608-624`). The current Rust harness and
driver exercise only `GeneratedParagraphSource`, whose provider is the
deterministic replayable bounded cursor (`docx_replayable_tail_append.rs:471-598`,
`:2140-2145`). There is no one-shot producer, explicit replay-store custody,
retained-store byte accounting, or store/storage arm in `measure.py`.

That is a blocker for the full M2/design measurement claim. The deterministic
ten-case subset can be captured and reviewed as an isolated arm, but it cannot
be presented as evidence for one-shot bounded-store behavior until that arm is
implemented and measured, or the design scope is explicitly revised with a
separate acceptance decision.

The same design section calls for owned, filesystem-backed, instrumented
short-read/range, and configurable high-latency input arms, several sequential
sink write sizes, and separate Store/Deflate selected-member cases. The current
driver fixes one caller-owned in-memory positional source contract and one 4 KiB
short sink (`measure.py:45-69`, `:214-224`); the generated package uses its one
default package-writer framing (`docx_replayable_tail_append.rs:890-904`). Those
missing arms must be recorded as out of scope for a deterministic subset or
closed before a broader end-to-end claim.

## Required disposition before formal results

Formal captures should remain **pending** until all of the following are true:

1. A frozen protocol and matching release normal/allocator build receipts exist,
   with unchanged source manifests and retained copied binaries.
2. All 40 formal process receipts pass, each has 30 measured samples and three
   warmups, and the frozen validator passes every per-sample check, including
   five authored opens, histogram conservation, sink proof equality, and
   allocator invariants. The frozen gate now rejects nonzero failed allocations,
   forged histogram ranges/largest buckets, wrong exact near-limit text size,
   and a mismatched pinned-target workspace formula. Any missing procfs sample
   must still be surfaced as a stated limitation, and omitted source/candidate/
   output ceilings must not be claimed as separately reported dimensions.
3. Cross-role and repeat equality of deterministic source/authored/oracle
   hashes is checked, and GNU `time` RSS is kept separate from operation RSS,
   VmHWM endpoints, and allocator region increments.
4. The missing one-shot store arm and any required input/sink/package-format
   arms are either captured under the same custody rules or explicitly removed
   from the claimed scope by the design owner.

Until then, the dev46 tests establish focused harness correctness and dev50
establishes a diagnostic execution path only. Neither is a formal performance
measurement.
