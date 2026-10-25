# 0362ul89arc — source admission and comment date-time limits

Date: 2026-10-25

This follow-up addresses the complete audit of `0362ul89arb`.

## Source repair

Source commit: `6601d91d70` (`fix(pptx): admit comment source attributes before
MCE`).  It is based on the 89arb full tip
`ca8c250590865d974a38deb7ac6cf8aabf02d27e`.

The ordinary legacy-comment parser now performs a borrowed raw attribute
admission before `process_comment_mce`.  An oversized attribute in an inactive
AlternateContent branch therefore remains physically refused instead of
vanishing during MCE.  The regression test checks the typed
`XmlAttributeBytes` refusal and the existing selected semantic aggregate still
applies after MCE.

Collaboration replay now constructs `MceSource` once from the already completed
raw scan, then reuses that filtered byte view for semantic MCE processing and
offset translation.  The extra `without_processing_instructions` source scan
and copy are gone.  The production selected parser remains a single fused
bounded event/model pass; the exact-boundary counter remains test-only.

`validate_comment_list_with_limit` now applies the selected `max_string_bytes`
to `date_time` before syntax validation, matching read and writer behavior.
The direct slide-comment writer has a regression test for an over-budget valid
date-time value.

## Evidence fence and checks

The retained external diagnosis log is
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ef9470f7da/fullfeatures-tests/0000.log`
(548,761 bytes, SHA-256
`cfe5aa3f9c4082d67019cfc8285a121768d0a5e3063f8e0c064bfcbda60d3fef`).

No Cargo, build, rustdoc, benchmark, or qualification command was run.  The
edited sources passed rustfmt parsing checks and `git diff --check`; repository
evidence-size and binary guards were clean before commit.  Final composed
compile, runtime qualification, and performance review remain coordinator
work.
