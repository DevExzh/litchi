# 0782 archived borrowed-slice candidate

This archive is a frozen candidate against base commit
`f1df64d15a8f2711e554f58acb413b06d56ef1ca`, whose PPT production source is
restored after the rejected 0781 `Cow<str>` experiment. The candidate is
archive-only and is not applied to the live crate. `candidate/model.patch` is
relative to that base, while `candidate/files/` contains complete replacement
files for every path in the patch. No Cargo, compiler, formatter, native
probe, profiler, or performance command was run for this archive.

The design follows the accepted API-layer, crate-topology, immutable-source,
memory/performance, validation/compatibility, migration/verification, and
current-topology constraints in ADRs 0001, 0002, 0003, 0005, 0006, 0008, and
0024. The public writer API and the `ShapeProperties.text: Option<String>`
owner remain unchanged; this is a crate-private serialization boundary.

## Narrow ownership change

`UserShapeData<'text>::text` is `Option<&'text str>`. Plain fresh conversion
borrows `props.text` with `as_deref()` while the source `WritableShape` remains
alive through synchronous serialization. The shape encoder passes the copied
optional slice to the existing `&str` textbox writer. No `Cow` variant or
owned plain-text alternative is retained because the current production call
graph never detaches a `UserShapeData` with non-empty plain text.

Rich paragraphs remain owned: conversion clones them before assigning PP9
smart-tag run identifiers. The centered/non-left plain fallback remains the
existing owned `Paragraph::new(text.clone()).align(...)` path. Empty text,
text interactions, notes defaults, refusal ordering, and all other shape
fields retain their existing routes.

The lifetime is propagated through `UserShapeData`, `ChildShape`,
`GroupShape`, `GroupChild`, both package vectors, drawing helpers, group
helpers, shape codecs, and validation. These types are crate-private and no
new dependency, unsafe code, cache, public source generic, or existing-document
policy is introduced.

## Changed-file allowlist

`candidate/changed-files.json` is sorted and contains the complete 12-file
crate source allowlist required by the source census. The nine production
paths are:

- `crates/litchi-ppt/src/writer/core/codec.rs`
- `crates/litchi-ppt/src/writer/core/package.rs`
- `crates/litchi-ppt/src/writer/escher/codec/drawing.rs`
- `crates/litchi-ppt/src/writer/escher/codec/group.rs`
- `crates/litchi-ppt/src/writer/escher/codec/properties.rs`
- `crates/litchi-ppt/src/writer/escher/codec/shapes.rs`
- `crates/litchi-ppt/src/writer/escher/codec/validation.rs`
- `crates/litchi-ppt/src/writer/escher/model.rs`
- `crates/litchi-ppt/src/writer/escher/semantic.rs`

The three focused test paths are:

- `crates/litchi-ppt/src/writer/core/tests.rs`
- `crates/litchi-ppt/src/writer/escher/tests/groups.rs`
- `crates/litchi-ppt/src/writer/escher/tests/legacy.rs`

`crates/litchi-ppt/src/writer/escher/codec/text.rs` is unchanged. The existing
`crates/litchi-ppt/tests/ppt_writer_text_goldens.rs` suite remains the broad
byte-equivalence gate for ASCII, Unicode, empty, record-boundary,
interaction, rich-text, long-text, notes, and both write/save routes; it is
intentionally not copied into this candidate because its source is unchanged.

## Focused gates

- `plain_text_conversion_borrows_source_string` compares the converted
  `&str` pointer and length with the source `String`.
- `borrowed_text_option_uses_slice_layout` checks the current target's safe
  `Option<&str>` layout against its slice and owned-string counterparts using
  `size_of`; it uses no unsafe code and does not assert a whole-struct layout.
- `test_plain_text_alignment_and_rotation_reach_escher_shape` keeps centered
  text in owned paragraphs and checks the retained run text.
- `smart_tags_round_trip_through_both_output_paths` verifies conversion does
  not leave PP9 identifiers in source rich-text runs.
- `grouped_shape_accepts_borrowed_text` passes a source-backed group shape,
  checks the decoded text, and compares the nested textbox payload bytes with
  the direct existing `build_client_textbox` output.
- Existing Escher tests and the complete PPT text golden suite remain required
  for byte, semantic, and refusal parity after application.

The layout assertion and the narrower representation are hypotheses for the
next paired 0782 matrix. They do not attribute the rejected 0781 short-text
latency regression or predict a performance result. Root should apply this
patch only after the frozen 0782 baseline qualification, then format, compile,
run focused and full byte/semantic gates, and perform the same paired native
and allocation measurements.

## Archive integrity

The following SHA-256 values cover every complete archived source file and the
two candidate control artifacts. The machine-readable receipt also records
the implementation-notes hash after this file is frozen.

| archived path | bytes | SHA-256 |
| --- | ---: | --- |
| `files/crates/litchi-ppt/src/writer/core/codec.rs` | 20836 | `7f453675adf9258303af5f8e8a0ae1fe4f5a5699229798127e2bbde9b82817fa` |
| `files/crates/litchi-ppt/src/writer/core/package.rs` | 51587 | `badb166ee253c80d5bfa3b2faf7fb0f86e317e5deeb19d7fbe75dcecae033de4` |
| `files/crates/litchi-ppt/src/writer/core/tests.rs` | 28327 | `74f3e08f7128dd58f27b9386749bfd9078dda3db219d01d10cd81b8d89a018b2` |
| `files/crates/litchi-ppt/src/writer/escher/codec/drawing.rs` | 9587 | `ba8ef8192bf1992fa20b7a9b0555deaa6f1d5f6c5d2fe61aa753fd2392735642` |
| `files/crates/litchi-ppt/src/writer/escher/codec/group.rs` | 4228 | `7882bb79da431d61efade55368edc97ceb4df7e1596e46a91c72dda0d3a35e6b` |
| `files/crates/litchi-ppt/src/writer/escher/codec/properties.rs` | 5855 | `f8891e0b1616cb1c6c9e2dd2de80ad83221481e2f4547f9bda80fc5b1c97ff24` |
| `files/crates/litchi-ppt/src/writer/escher/codec/shapes.rs` | 10740 | `28abfdec9e65491fe53274fa0b4c7a7d6ecd65afc99898c51ede4c3b75c96b2f` |
| `files/crates/litchi-ppt/src/writer/escher/codec/validation.rs` | 2838 | `b82828219643d848ea83d3814bedf6c3404233900c384f4943fb865505f39e22` |
| `files/crates/litchi-ppt/src/writer/escher/model.rs` | 20411 | `fbe3723b4dbae425664f03b2e2ddd10ed9b5f183948472b2bb432a3c06989d90` |
| `files/crates/litchi-ppt/src/writer/escher/semantic.rs` | 3299 | `26941995cd59539bc024ab722901c53e91af7365c099a8b9ec40f56f1d956700` |
| `files/crates/litchi-ppt/src/writer/escher/tests/groups.rs` | 5016 | `d95a3ba9825526b28fa5a7b4766637e0072e24d0094cb31417f735d6d2d590e3` |
| `files/crates/litchi-ppt/src/writer/escher/tests/legacy.rs` | 29137 | `079ddcf258f502f926a31cc8a7995bfbe8f1162a8594f742e618a0ae4d714520` |
| `model.patch` | 18824 | `35a7af255660310db93214333b2b58247bcf2cf85fc9b6a85bfa3c4476d9f4bf` |
| `changed-files.json` | 659 | `6cd27fd62fbd54478ff8bdcde7d9ecc548c0b597f29e31d3c5baab3a2a381c81` |
