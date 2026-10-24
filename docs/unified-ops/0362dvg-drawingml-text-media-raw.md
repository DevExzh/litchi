# 0362dvg — DrawingML text and media raw limits

- Date: 2026-10-24.
- Parent: 0351, 0362dv, 0362dvb.
- Status: source migration complete; coordinator validation and integration remain.

This lane gives the remaining shared DrawingML raw codecs one required policy
at every resource-scaled boundary. Text-body reads now use the single
`reader::read(xml, scope, limits)` entry; text-body writes use
`writer::{write,write_contents}(..., writer::Limits)` and count the complete
output before reserving or changing the caller's string. Their existing
`Builder` and fixed-domain scalar APIs remain unchanged.

`model3d::codec::{read,opaque,write,write_to}` now take
`model3d::codec::Limits`. The read policy bounds complete input, opaque
fragments, child counts, raster children, depth, and nested nodes. Writers
preflight the serialized size; `write_to` then emits directly to the caller's
sink rather than materializing a second complete buffer. Invalid policies use
`InvalidLimit`, and over-budget failures retain the typed `LimitKind` plus
observed and maximum values.

SVG blips now expose one top-level API with `svg_blip::Limits`; the former
nested codec is crate-private. Standalone and contextual reads, vector writes,
and both sink writes all require the policy. Contextual output performs the
existing size preflight before allocation or the first sink write. The
arbitrary-XML `blip::find_first_embed` scanner likewise requires its typed
input/relationship policy and checks input admission before constructing the
reader.

PPTX, DOCX, XLSX, spreadsheet-drawing, DrawingML tests, and the live XLSX SVG
allocation harness were migrated in this source change. Historical baseline
and matched evidence snapshots under `docs/report/spec-gap-validation-evidence`
were left byte-for-byte unchanged. The glossary now bans the removed
`text::body::reader::read_with_options` spelling in favor of the canonical
required-limits entry.

No Cargo command was run in this worker before the coordinator grants the
validation slot. `rustfmt`, `git diff --check`, and glossary JSON parsing were
run locally. No generated binaries, third-party paths, or large evidence
artifacts are part of this commit.
