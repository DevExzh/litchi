# 0362ul89aga — PPTX physical duplicate-attribute limit repair

Date: 2026-10-25

## Review finding

The first 89ag repair restored the established `first_wins` semantic readers
in custom-show and handout metadata, but its admission count also skipped
duplicate names. That allowed a tag with a selected limit of two physical
attributes to carry two or more duplicate occurrences after one retained name.
It also left the existing 40,001-attribute first-wins stress fixtures above
the old 20,000 default if physical occurrences were counted.

## Repair

Commit `5e18c26042921688262dbdb4a764b350f20cad49`
(`fix(pptx): count duplicate metadata attributes physically`) updates only
the metadata limits and the custom-show/handout codecs:

- `count_up_to(element, limits.max_attributes())` now preflights every
  physical attribute occurrence, including duplicates, before `first_wins`.
  It is bounded linear work and reports the exact `maximum + 1` admission
  observation for an over-limit tag.
- `first_wins` remains the semantic parser. Duplicate errors are accepted,
  non-duplicate lexical errors remain `Error::Xml`, and first values still win.
- `Limits::DEFAULT.max_attributes` is 65,536. This is the smallest practical
  rounded ceiling above the unchanged 40,001-attribute stress fixtures while
  keeping the physical scan bounded. `Limits::new`, `with_attributes`,
  `from_source_limits`, and source/MCE admission continue to derive or
  intersect their maxima from the selected profile; no tighter profile is
  widened.
- Both codecs now test an exact three-attribute duplicate tag and a selected
  two-attribute one-under case. The crowded stress fixtures were not reduced.

The source patch is retained at:

`(deleted-coord-artifact)/metadata-runtime-89aga-physical-limit-5e18c26042.patch`

It is 8,229 bytes and has SHA-256
`640291e07b1eb60d5d12c31cfc397ddc12323ee6c68d5e39a12329f72b8cb4ed`.

## Evidence and status

The triggering actual-runtime capture remains the 89ag evidence against
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`:

`(deleted-coord-artifact)/0000.log`

It is 378,736 bytes with SHA-256
`4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a`.

No Cargo command, build, rustdoc, test, benchmark, or Clippy run was made for
this follow-up. Source-only parsing passed for all three changed Rust files
with `rustfmt --edition 2024 --config skip_children=true --emit stdout`, and
`git diff --check` passed. Runtime and performance qualification remain
pending coordinator composition and a fresh isolated run.
