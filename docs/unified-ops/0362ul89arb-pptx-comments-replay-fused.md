# 0362ul89arb — collaboration replay admission ordering and fusion

Date: 2026-10-25

This is the follow-up to `0362ul89ara`, whose composed review found that the
semantic MCE pass ran before raw physical admission and that the new selected
preflight added a successful-path scan and a second baseline MCE pass.

## Source repair

Source commit: `c4bfeeec18` (`fix(pptx): fuse selected comment admission`).
The source patch is based on the 89ara full tip
`5b9ada193fc4bb68a7ac36afbad0a74f6a53c3f1`.

`locate` and replay validation now run the source-bound raw scan first.  The
scan checks input, depth, node/event, duplicate/namespace structure, and the
independent raw `max_attribute_bytes` ceiling before MCE can allocate a
processed projection.  The existing scan's borrowed start/empty-event check
remains ahead of its `into_owned` event conversion, and the ordinary comment
parser checks the same physical ceiling before decoding attribute values.

The selected parser now has an internal processed entry point.  It does not
run baseline MCE again, and its existing event pass counts decoded aggregate
string bytes before owning an attribute or text value.  Character/entity
lengths are precomputed without allocation; `dt` receives an explicit selected
string check alongside author names/initials and comment text.  Thus the
successful path retains one raw source scan, one selected MCE projection, and
one selected grammar/model pass; the former standalone selected preflight and
second baseline MCE pass are gone.  The test-only boundary counter remains
under `cfg(test)` for exact/one-under regressions and is not a production pass.

Strict threading replay still validates only the MCE-selected projection, so
inactive zero-index/oversized ordinary fields remain semantically opaque while
the raw source continues to pay physical limits.

## Evidence fence and checks

The retained external diagnosis log is
`(deleted-coord-artifact)/0000.log`
(548,761 bytes, SHA-256
`cfe5aa3f9c4082d67019cfc8285a121768d0a5e3063f8e0c064bfcbda60d3fef`).

No Cargo, build, rustdoc, benchmark, or qualification command was run.  The
hydrated source files passed rustfmt parsing checks and `git diff --check`; the
repository evidence-size and binary guards were clean before commit.  The
coordinator must perform the composed compile and runtime qualification.
