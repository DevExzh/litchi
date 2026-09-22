# 0734 candidate implementation notes

This directory contains a candidate copy of the PPT embedded-editor finish
module. The live crate is intentionally unchanged while the root agent runs
the ordinary baseline and decides whether to apply the candidate.

The candidate consumes the private `Editor` at the existing `finish` boundary,
then transfers the already-owned `Editor::streams` payload vectors into the
existing `OleWriter::create_stream_owned` API. The newly assembled incremental
Document stream and updated Current User stream take the same route. Source
layout adoption still happens before stream registration, stream insertion
order is the order returned by `Editor::streams`, the selected sector policy
is unchanged, and the bounded output sink plus final `validate_rewrite` remain
in place.

`Editor::open` obtains `streams` from one `OleFile::list_streams` result. The
document and Current User paths are selected by the first matching leaf name;
therefore the candidate does not assume that a leaf name is globally unique.
For the normal topology, each selected complete path occurs once and its
payload is moved exactly once. If a synthetic or malformed editor contains a
duplicate selected complete path, those selected paths use the old borrowed
`create_stream` behavior, preserving replacement and allocation-error
semantics for that reachable case. Missing selected paths retain the old
behavior: no stream is injected and the assembled payload is dropped after
the writer call. Auxiliary paths are moved individually in every case.

The only new `Corrupted` branches guard an internal ownership contradiction
after the occurrence count has selected the one-shot path. They are not
reachable from the counted loop, and prevent an accidental panic if this
candidate is changed later. They do not replace any parser, package, source,
mapping, output-limit, Reuse-plan, final-reopen, or public semantic
validation.

The candidate tests are kept in the copied `finish.rs` so the root can apply
one source file after the baseline. They compare the candidate writer with a
test-only copy of the previous borrowed writer for both `Reuse` and `Rewrite`,
exercise a real incremental record replacement and reopen, assert exact
no-op/source immutability, and retain finite output-limit coverage. A test
helper is deliberately test-only; production code has no second writer path.

No Cargo, native benchmark, formatter, or live-crate command is run by this
candidate owner.
