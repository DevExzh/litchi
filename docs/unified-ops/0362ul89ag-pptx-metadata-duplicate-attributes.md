# 0362ul89ag — PPTX metadata duplicate-attribute admission repair

**Returned proposal:** this semantic-only attribute-count design was not
accepted. Record 89aga replaces it with a physical count including duplicates
and a 65,536 default ceiling; the 20,000-name rationale below is historical.

Date: 2026-10-25

## Failure diagnosis

The isolated 0362ul89 actual-runtime capture was run against source commit
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`. Its full-feature library row
reported 1,356 passes, 23 failures, and 1 ignored test. Four failures were in
the metadata scope:

- `custom_show::codec::tests::crowded_custom_show_attributes_are_read_first_wins`
- `custom_show::codec::tests::parsed_custom_show_attributes_keep_their_first_occurrence`
- `handout::codec::tests::crowded_handout_attributes_are_read_first_wins`
- `handout::codec::tests::parsed_handout_attributes_keep_their_first_occurrence`

Both codecs already used `first_wins` for semantic reads, but their new
resource-admission helper used `checked_attributes`. The checked iterator
reported a duplicate as `Xml` before the established first-wins readers could
consume the element. The crowded tests intentionally exercise 20,000 unique
attribute names followed by duplicate names, so counting every physical
duplicate would also incorrectly exceed the 20,000-name semantic attribute
budget.

The retained external failure evidence is:

`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ae6f32d202/fullfeatures-tests/0000.log`

It is 378,736 bytes with SHA-256
`4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a`.

## Repair

Commit `9c371c0230263d933618edeec1892f05d424f775`
(`fix(pptx): admit duplicate metadata attributes`) changes only
`custom_show/codec.rs` and `handout/codec.rs`. Admission now iterates the same
`first_wins` stream as the model readers, counts only first occurrences toward
`max_attributes`, accepts `AttrError::Duplicated`, and preserves `Error::Xml`
for other lexical attribute errors. The selected string and attribute checks
still apply to every retained first occurrence; raw duplicate bytes remain
bounded by the existing input/output limits. No shared helper, model, test,
or capability was changed.

An external patch against the exact actual-runtime base is retained at:

`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89/metadata-runtime-89ag-9c371c0230.patch`

It is 2,767 bytes and has SHA-256
`aee1fbe36766dc920438c66df3d1e47811505bd1559a7b14ce3d0ed6069bae2a`.

## Checks and status

No Cargo command, build, rustdoc, test, benchmark, or Clippy run was made for
this repair. Source-only parsing passed for both changed Rust files with
`rustfmt --edition 2024 --config skip_children=true --emit stdout`, and
`git diff --check` passed.

The capture remains evidence of the pre-repair state. A fresh composed
actual-runtime run is required before acceptance.
