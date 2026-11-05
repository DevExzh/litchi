# 0362ul89ara — bounded selected collaboration replay

Date: 2026-10-25

## Scope and source fence

This leaf repairs the returned `89ar` collaboration replay change on top of
`e7c19964b28085753a8c9ff2955225e796098a73`.  The source commit is
`abc8874b84` plus follow-up `8d2866c8f8` (`fix(pptx): check raw comment
attributes before ownership`).  The first record commit is `41bece51cb`; the
final worktree tip also includes the small record update below.
It is an external patch for the coordinator's current composed candidate; this
worktree does not modify the candidate, merge branches, or `PROGRESS.md`.

The qualification diagnosis remains the external log
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ef9470f7da/fullfeatures-tests/0000.log`
(548,761 bytes, SHA-256
`cfe5aa3f9c4082d67019cfc8285a121768d0a5e3063f8e0c064bfcbda60d3fef`).
No Cargo, build, rustdoc, native, benchmark, or qualification command was run
in this repair lane.

## Returned-review repairs

The earlier replay validator raised its parser string ceiling to the input
ceiling and then checked selected fields after the ordinary parser had already
owned them.  This leaf replaces that widening with a borrowed quick-XML
preflight over the MCE-selected projection.  It charges the same aggregate
decoded attribute-name/value and text policy as the ordinary parser, and checks
author `name`/`initials`, comment text, and comment `dt` before model-owned
strings are built.  Entity lengths are counted without constructing decoded
values, so exact and one-under aggregate boundaries remain observable.

`validate_threading_replay` now performs strict ordinary-comment validation on
the selected projection, rather than validating the raw source before MCE.
Inactive AlternateContent branches therefore remain semantically opaque while
their raw bytes still pass through the physical source guards.

`comments::Limits` now carries `max_attribute_bytes` independently from the
selected `max_string_bytes` policy.  Both the ordinary comment parser and the
source-bound collaboration scan reject an over-sized raw attribute value before
decoding or owning it.  The selected semantic budget may therefore be tighter
than the physical ceiling without charging inactive presence or threading
values; source-constrained profiles reduce both ceilings together.

The follow-up also checks borrowed raw start/empty events before the
source-bound scan calls `into_owned`, and the ordinary parser now uses the
borrowed slice reader.  This keeps the physical attribute admission ahead of
decoded attribute ownership on both related paths.

The source tests cover exact/one-under selected aggregate admission, active
ordinary-field growth, `dt`, inactive presence and threading values, strict
post-MCE threading, and an independent physical attribute ceiling.

## Source-only checks

`rustfmt --emit stdout` parsing checks and `git diff --check` passed.  The
repository evidence-size and binary guards were clean before the source commit.
The coordinator must run the composed Cargo qualification and review the
resulting capability/performance evidence before accepting this leaf.
