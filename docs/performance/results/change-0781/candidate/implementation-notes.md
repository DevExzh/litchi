# 0781 archived candidate

This archive is a frozen candidate against base commit
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8`. It is not applied to the live
crate. `candidate/model.patch` is the baseline-relative patch, and
`candidate/files/` contains complete replacement files for every path in the
patch. The patch was checked with `git apply --check`; this agent did not run
Cargo, a compiler, a formatter, or a performance probe.

## Ownership design

`UserShapeData<'text>::text` changes from `Option<String>` to
`Option<Cow<'text, str>>`. The public writer model remains
`ShapeProperties.text: Option<String>`, so this is an internal conversion
boundary only. The conversion functions take `&'text WritableShape` and use
`Cow::Borrowed` for the plain-text branch when the source has no paragraphs.
The Escher text encoder consumes `&str` through `Cow::as_ref()`, preserving
the existing encoding path and output bytes.

The borrow is valid for the synchronous fresh serialization call. The
intermediate `UserShapeData` values do not outlive the source writer shapes.
`ChildShape<'text>`, `GroupShape<'text>`, and `GroupChild<'text>` carry the
same lifetime so grouped shape inputs can retain borrowed text. Their codec
and validation helpers use the corresponding lifetime-parameterized types;
there is no public API or dependency change.

Rich paragraphs stay owned. Conversion clones them because smart-tag mapping
assigns PP9 run identifiers to the converted copy. The centered/non-left
plain fallback also stays owned through `Paragraph::new(text.clone())`,
which preserves its paragraph and alignment behavior. No unsafe code,
global cache, or existing-document policy is introduced.

## Archived file allowlist

`candidate/changed-files.json` is the sorted complete crate allowlist for the
source census and includes all 12 files carried by `model.patch`. Its
production paths are:

- `crates/litchi-ppt/src/writer/core/codec.rs`
- `crates/litchi-ppt/src/writer/core/package.rs`
- `crates/litchi-ppt/src/writer/escher/codec/drawing.rs`
- `crates/litchi-ppt/src/writer/escher/codec/group.rs`
- `crates/litchi-ppt/src/writer/escher/codec/properties.rs`
- `crates/litchi-ppt/src/writer/escher/codec/shapes.rs`
- `crates/litchi-ppt/src/writer/escher/codec/validation.rs`
- `crates/litchi-ppt/src/writer/escher/model.rs`
- `crates/litchi-ppt/src/writer/escher/semantic.rs`

The allowlist also carries these focused test changes:

- `crates/litchi-ppt/src/writer/core/tests.rs`
- `crates/litchi-ppt/src/writer/escher/tests/groups.rs`
- `crates/litchi-ppt/src/writer/escher/tests/legacy.rs`

`crates/litchi-ppt/src/writer/escher/codec/text.rs` is intentionally
unchanged. The existing `crates/litchi-ppt/tests/ppt_writer_text_goldens.rs`
suite remains the byte-equivalence gate for ordinary, Unicode, empty,
record-boundary, rich-text, and long-text cases; it is not duplicated in the
archive.

## Focused checks carried by the candidate

- `plain_text_conversion_borrows_source_string` checks `Cow::Borrowed` and
  compares the converted slice pointer and length with the source `String`.
- `test_plain_text_alignment_and_rotation_reach_escher_shape` keeps the
  centered fallback on owned paragraphs and checks the paragraph text.
- `smart_tags_round_trip_through_both_output_paths` verifies conversion does
  not leave PP9 identifiers in the source runs after mutating the copied rich
  paragraphs.
- `grouped_shape_accepts_borrowed_text` encodes a grouped borrowed shape and
  reads the text back from the Escher textbox.

After baseline qualification, the coordinator should apply this patch,
format and compile it, run these focused tests plus the complete PPT text
goldens, and then run the paired quality and performance gates from the
0781 plan. Adoption still depends on exact output/refusal parity and fresh
paired measurements; the historical 0753 attribution is motivation only.

## Archive integrity

The following SHA-256 hashes cover every complete archived source file and
the two candidate control artifacts. Byte counts are included to make partial
or stale copies visible.

| archived path | bytes | SHA-256 |
| --- | ---: | --- |
| `files/crates/litchi-ppt/src/writer/core/codec.rs` | 20877 | `501324a7c17f5a2919961f4883533555f6c6b1d9904a8b5a9b8fe2a80a3cc767` |
| `files/crates/litchi-ppt/src/writer/core/package.rs` | 51587 | `badb166ee253c80d5bfa3b2faf7fb0f86e317e5deeb19d7fbe75dcecae033de4` |
| `files/crates/litchi-ppt/src/writer/core/tests.rs` | 28255 | `2ebe6c8045c19198699a792bb2c57cc9003ff035fe3e54c1d0633b32eb9afbc4` |
| `files/crates/litchi-ppt/src/writer/escher/codec/drawing.rs` | 9587 | `ba8ef8192bf1992fa20b7a9b0555deaa6f1d5f6c5d2fe61aa753fd2392735642` |
| `files/crates/litchi-ppt/src/writer/escher/codec/group.rs` | 4228 | `7882bb79da431d61efade55368edc97ceb4df7e1596e46a91c72dda0d3a35e6b` |
| `files/crates/litchi-ppt/src/writer/escher/codec/properties.rs` | 5855 | `f8891e0b1616cb1c6c9e2dd2de80ad83221481e2f4547f9bda80fc5b1c97ff24` |
| `files/crates/litchi-ppt/src/writer/escher/codec/shapes.rs` | 10750 | `0e79993e04dcaed88220e65cd23255ea82fdbcd1ff07604733499856d2e0d083` |
| `files/crates/litchi-ppt/src/writer/escher/codec/validation.rs` | 2853 | `a82f4e928bb50d28f8d88f40587f624f15b0438586b4d1493c5b39975164dfed` |
| `files/crates/litchi-ppt/src/writer/escher/model.rs` | 20438 | `b22ff0302549dc2cccd5bdd589b5398fd76b836f0b2e46a3b11903d9bf4a808e` |
| `files/crates/litchi-ppt/src/writer/escher/semantic.rs` | 3377 | `13739f9cc5a6ee0966514118a2c4acf3dc45099749e0e66b61493a3c4af6628c` |
| `files/crates/litchi-ppt/src/writer/escher/tests/groups.rs` | 4914 | `1324c79a68f9c7aa2d3282df3a81863fc88ef7c21b3f19f0ae9fedbe7f71b0bb` |
| `files/crates/litchi-ppt/src/writer/escher/tests/legacy.rs` | 29144 | `97aea4eb57de97e5928c661130bcd3d494da4fed3d70958dba12f8089b37a5e9` |
| `model.patch` | 19364 | `329672c7300e08c659f0716572f65e038bbeefddb2e5e51eb813cb530f87fc3a` |
| `changed-files.json` | 659 | `6cd27fd62fbd54478ff8bdcde7d9ecc548c0b597f29e31d3c5baab3a2a381c81` |
