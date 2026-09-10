# Next implementation after 0496

0496 provides descriptive phase evidence for the opened-document DOCX
lifecycle. Its formal analysis and verification are retained in
[`analysis/formal1.json`](analysis/formal1.json) and
[`verification/formal1.json`](verification/formal1.json). The result does not
authorize a production optimization: the historical 0495 flags remain
unresolved, and the phase percentages are wall-clock observations from one
serial lifecycle.

The publication interval is also a mixed measurement boundary. In the 0496
harness, the timed publication writes the complete candidate through a
preallocated `Vec` sink, so it includes the output byte copy. The output SHA,
semantic/media verification, and other preservation oracles run after the
timed interval. Those costs must remain mandatory evidence, but they must be
reported separately from production publication physics before a publication
optimization is selected. The relevant boundary is in
[`docx_managed_edit.rs`](../../../../tools/perf-baseline/src/docx_managed_edit.rs),
where `SequentialSink` accepts the timed writes and the output checks follow
the lifecycle clock.

The next production change should close the missing atomic destination route
for the existing bounded DOCX logical-tail append. The route already has
finite source, authored, candidate, replay-window, work, cancellation, source
version, and artifact-fingerprint contracts in
[`tail_append_stream.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs).
It currently exposes sequential `write_to_stream` publication through
`ParagraphStreamPlan` and `ParagraphStreamCommit`, but no path-destination
publication. Add a consuming `write_to_path` method at those owners that
publishes the prepared splice into a sibling temporary file and calls the
existing typed atomic primitive
[`replace_with`](../../../../crates/litchi-opc/src/atomic.rs). The method must
return the same `ParagraphStreamPublication` proof product after replacement;
it must not weaken preparation or source-freshness checks or introduce a new
adapter-only route.

The path contract is:

- validate the destination before replacement, reject symlinks and
  non-regular files, preserve existing permissions, write and `sync_all` the
  sibling temporary file, persist it, and sync the parent directory;
- leave the destination unchanged when preparation, source validation,
  cancellation, limits, temporary-file writing, or synchronization fails
  before replacement; propagate `OpcError::Committed` when replacement
  happened but the parent-directory sync failed;
- retain the candidate artifact fingerprint and length, source fingerprint and
  length, authored proof, replay reference, durable patch, and exact inverse
  authorization returned by the current stream publication;
- keep managed output, memory, object, work, and replay-window accounting
  balanced on success and every failure path.

Add focused route tests in
[`source_backed_tail_append_stream.rs`](../../../../crates/litchi-docx/tests/source_backed_tail_append_stream.rs):
successful new and existing destinations with reopen and exact candidate
oracles; injected sink/limit/cancellation failures with an unchanged
destination and no temporary-file residue; stale, foreign, and mutated-source
refusal before destination mutation; symlink/non-regular destination refusal;
durable patch and inverse replay after reopening the path; and managed budget
release. The generic atomic tests remain useful, but they do not prove that
the DOCX stream's source proof, replay store, and publication proof are wired
through the atomic route.

Measure the route in separate scopes. Preserve the existing sequential
`write_to_stream` case, add a counting/non-retaining sink arm to isolate
publication byte-copy and write-call cost, and keep SHA-256, semantic/media,
reopen, patch, and inverse checks outside the timed interval as mandatory
oracles. Add a separately labelled atomic-path case whose interval includes
temporary-file writes, `sync_all`, replacement, and parent-directory sync.
Use the existing source/authored matrix (owned, `FileSource`, short-read,
delayed provider; deterministic, memory-store, and file-store authored input),
record source and replay counters, source before/after identity, accepted
output bytes and write distribution, candidate digest, and destination
before/after state. Run normal and allocator roles with reversed repeats and
bounded source/authored sizes. Genuine borrowed input, native Office
producers, filesystem-cold intersections, and concurrent workers remain
separate unfulfilled cases.

Bounded concurrency should follow this route. The current DOCX lifecycle is
one-worker and serial, and the repository has no independent work partition
or end-to-end 1/2/4/8-worker contract for this operation. Adding workers
before defining that partition, backpressure budget, deterministic output
oracle, and Amdahl measurement would be a larger architectural change rather
than a measured follow-up.

This change closes one concrete filesystem-atomic-save and bounded logical
append gap. The broader non-iWork goals remain open, including genuine
borrowed lifetimes, native producer round trips, cold/provider intersections,
bounded parallel scaling, durable history and composition, and the remaining
CRUD, conversion, structural, dynamic-content, security, malformed-input, and
format coverage.
