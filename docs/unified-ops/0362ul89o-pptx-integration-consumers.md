# 0362ul89o — PPTX integration-consumer call-shape repair

Date: 2026-10-25. This is a source-only integration-test/example repair; no
compiler, runtime, performance, or qualification result is claimed here.

The source base is `2749ced58b3837c477fad3a77bed633d922cc252`. The preserved
F5 diagnostic log is outside Git at
`(deleted-coord-artifact)/01-pptx-all-features-check-attempt-actual-f5-01.log`
(26,413 bytes, SHA-256
`6426504ce368cb6721a0c2cbcdc2161854bb85baf667a11336e6f3aaeb76b730`). Its
diagnostics index is outside Git at
`(deleted-coord-artifact)/diagnostics-index.json`
(363 bytes, SHA-256
`c4ef106aab2dbb61ac0ea2e07643c4b0019c9200f6d3bbadfad2f91a7a05820d`). The
index records exactly two errors, both missing the required second argument to
`transition::write` in `pptx_transition_extensions.rs` at lines 75 and 97.

Both fixture writer calls now pass the explicit
`litchi_pptx::transition::Limits::DEFAULT`. No production source, public alias,
default fallback, fixture, or return shape was changed.

The audit covered all 108 tracked integration-test Rust files and all 6
examples (28,079 Rust lines), including raw parser/writer calls and semantic
metadata owners migrated by the 0362ul61/62/63 lanes. Existing selected-limit
arguments were checked across animations, comments, media, notes, tags, laser,
model3d, shape, transition, view/presentation properties, and metadata
families. The two transition writer calls were the only stale required-limit
call sites found; existing high-level owner loaders retain their selected
profile behavior.

No Cargo, rustdoc, build, or runtime command was run. `git diff --check` and
the repository evidence/binary/path guards are required before commit.
