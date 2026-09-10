# 0497 publication harness review

This is a read-only review of the current DOCX publication harness and the
retained pilot evidence. I ran no Cargo command, build, benchmark capture, or
profile. I did not access the protected
`/home/zhuhe/code/litchi-spec-gaps` worktree. The only file changed for this
review is this document.

## Frozen custody and pilot inventory

The frozen protocol is
`c47939dcb7fb44fa1a8a3ffe9524f211ee1e671c5209ca6a4a5090d87268227d`. It
expects 288 formal children and 8,640 formal samples, plus a separate 72-child
pilot with 216 samples. The four retained build records all bind revision
`66af25e2fc3c208e6819fffa5fa29e102942bc8e`:

| build | binary SHA-256 | source manifest SHA-256 |
| --- | --- | --- |
| before/normal | `d0ab7f0c88f798b3ed5ffe2d6702512b086aa63784f8ccb3c1ce5dddabb505a` | `665bf079393067607a9c96076b64a4f9134132e68a5bba9a63901dde543b6477` |
| before/allocator | `0c36f947b61ddef96b202b8f4e98a0e32ff6a5c276173330925ffb7c26976d0d` | `665bf079393067607a9c96076b64a4f9134132e68a5bba9a63901dde543b6477` |
| after/normal | `c7ea684cce1dfd383a5b7dda6fd6f901eca153feb9cc7feab03ccf25b61371c9` | `52c1deca9ed37ccc2ae53a6db5ba4ec8eaa1bad50ff2105429b351d7f6cf62e3` |
| after/allocator | `5d9af0bbd3185d9044d8e91bc8bafe949bd9896556e6c9b7d700ef3a8eb13e45` | `52c1deca9ed37ccc2ae53a6db5ba4ec8eaa1bad50ff2105429b351d7f6cf62e3` |

The candidate manifest is
`f4b7438287bbd70524c8a3596648493c3d7238ce7c6e0f32152f5db049a7f49a`, and
the retained six-file patch is
`e903b7a8e30884d3f23e1d67802864ededf704251713fa6289d32c4c505cb5fe`. Each
manifest entry matches the current file bytes, and the six patch paths equal
the manifest allowlist. The pilot's 72 directory labels exactly equal the
frozen `pilot_runs` labels and ordinals. They cover the planned normal-role
arms: source counts 64 and 131072, authored counts 64 and 16384, fixed64
chunks, short text, all four deterministic input modes where applicable, both
store providers, and the planned default/counting/atomic routes.

An independent read-only audit of `captures/pilots/pilot1` found 72 terminal
receipts and 216 report samples. Every terminal exited zero without timeout,
termination, launch, validation, or missing-artifact errors. All 72 private
roots have a passing cleanup receipt with no remaining paths; every retained
artifact path, byte count, and SHA-256 matches its terminal receipt. Before
captures bind the before binary and manifest, and after captures bind the after
binary and manifest; all 72 bind the frozen protocol and driver.

## Report shape and timed boundary

The default route remains JSON-compatible: all 36 default pilot reports omit
`config.publication` and all 108 default samples omit `sample.publication`,
while retaining the historical hashing sink fields and digest. Counting and
atomic reports add the publication record only for their after-only routes.
The prior validator mismatch around the serialized `schema` and
`timing_scope` fields is resolved. The current validator requires both fields
and checks the route-specific values.

The focused Rust publication tests exercise counting and atomic publication
through deterministic, memory-store, and file-store providers, and exercise
the default route's omitted extension. The Python measurement tests cover the
publication schema, atomic null semantics, selected-arm configuration, and
failed-child custody. The profile tests cover report-path binding, warmup
destination ordering, sibling write/sync ordering, rename, and
parent-directory synchronization.

The timed boundary is coherent in the Rust harness. Atomic destination
directory creation occurs before `Instant::now()`. The timed block admits the
source, prepares the package and stream plan, performs the selected production
publication, and drops the publication and lifecycle owners before the elapsed
timestamp. For atomic publication this includes the production sibling write,
file data synchronization, replacement, and parent-directory synchronization.
Destination readback, ZIP/semantic/source/inverse checks, replay-file cleanup,
report serialization, and process/allocation endpoint snapshots occur after
the timestamp. The emitted timing scopes match this boundary:

* counting: `source_admission_prepare_sequential_sink_publication_drop`;
* atomic: `source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop`.

The pilot reports show these scopes 54 times each, with no alternate value.

## Artifact and oracle evidence

Counting samples retain accepted-byte, write-call, largest-write, and histogram
observations. All 54 counting samples have accepted bytes equal to the
candidate archive length and a null sink digest; the digest and length come
from the timed `ParagraphStreamPublication` artifact proof. Atomic samples
intentionally expose no sink write counts or digest: all five serialized sink
fields are JSON `null` in all 54 samples. No synthetic atomic counters were
observed.

The case-level oracle passed for all 72 pilot reports. Every candidate XML,
candidate semantic text, untouched member metadata, untouched raw ZIP member,
physical member order, opaque member, source-unchanged, and inverse flag is
true. The 54 atomic samples additionally read the actual destination after the
timer; each destination's byte count and SHA-256 equal the candidate oracle,
the complete post-timer oracle equals the case oracle, and destination and
private parent cleanup both pass. The three candidate identities exercised by
the pilot are retained in the raw reports (2,336; 51,186; and 344,330 bytes).

The inverse claim is deliberately narrower than a timed destination inverse
execution. `AtomicDestination::verify_and_cleanup` applies the post-timer XML,
semantic, raw-member, order, opaque-member, and source checks to bytes read
from the actual destination, but passes the fixture-owned source archive as
the inverse input. The emitted scope is
`untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted`.
The pilot supports that stated scope. It must not be rewritten as proof that
the timed destination was independently inverse-replayed.

## Failure custody, path binding, and allocator scope

The formal driver now preserves a failed child's private tree and records a
bounded hash inventory when exit, timeout, launch, or report validation fails
(`measure.py`'s `preserve_failure` path). Successful children still require an
empty private tree and remove it. Focused measurement tests cover the retained
failure inventory. The pilot contains only successful children, so this review
does not claim an observed failed-child receipt.

The current report validator binds each atomic destination to its reported
private parent and to the child's private `TMPDIR`, and requires the cleaned
paths to be absent. The profile parser and focused parser tests also bind raw
trace destinations to report paths and check sibling write/sync, rename, and
parent-directory-sync order. The retained `strace1` attempt failed before its
benchmark child launched because the host rejected the `fstatat` filter. The
`profile_v2.py` helper removes that unsupported filter, but no `strace2`
receipt exists yet; there is therefore no syscall-path evidence to report.

The pilot is intentionally normal-role only. The allocator binaries and
source bindings are present in the four-build custody, and the measurement
validator checks allocator conservation and operation-region peak fields, but
allocator publication samples remain part of the running formal lane. No
allocator result is inferred from this pilot.

## Findings and remaining boundary

I found no production-change blocker in the current pilot implementation or
its successful artifact proofs. The remaining formal capture and allocator
receipts are required before any formal latency, allocation, RSS, source-read,
replay, or route comparison result is promoted. The failed `strace1` attempt is
incomplete diagnostic evidence, and `strace2` must remain a separate bound
attempt if it is run.

One nonblocking validator-hardening point remains: the atomic sink validator
accepts an omitted subset of the five unavailable sink fields for compatibility
with older receipts, provided every present value is null. The current pilot
emits all five fields as null and passes the stricter intended semantics. If
the frozen current JSON shape must be fail-closed against serialization drift,
require the exact five-key null shape in the current protocol version rather
than accepting the older omitted shape. This does not fabricate an observation
and does not require a production-code change.
