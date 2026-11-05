# 0362ul69 — Formula raw codec policy

## Scope

This leaf closes the Formula raw-codec gap identified by the 0362 audit. It
covers the public MTEF and OMML parser/writer boundaries and every public
conversion helper in `litchi-formula`, plus the DOC, DOCX, ops, example and
fuzz consumers of those APIs. LaTeX remains a semantic AST input helper; its
output-side helper takes the selected destination codec policy without
pretending that LaTeX text itself is a wire format.

## API

`mtef::Limits` and `omml::Limits` are finite, caller-selectable policies with
typed resource kinds for input bytes, depth, work, retained items, string
bytes and output bytes. `MtefParser::new` and `OmmlParser::new` require the
selected policy. `MtefWriter::new`, `MtefWriter::with_options` and
`OmmlWriter::new` retain it for the writer lifetime. The public conversion
helpers require the source policy at every raw ingress and the destination
policy at every raw emission. There is no limit-less constructor or
`*_with_limits` alias.

Parser admission checks input size before constructing the XML reader or MTEF
binary state. Event/record work and item counts are charged before nested
context or object allocation; depth and individual strings are checked at
their boundaries. Writers preflight the AST before reserving output, charging
work/items/depth/string sizes and a conservative wire-size bound, then reserve
only within the selected output ceiling. This preserves the existing MTEF
record layout and OMML namespace/element spelling while preventing growing
unbounded output buffers.

## Consumer migration

The same commit migrates the Formula unit tests, examples, fuzz target, DOC
MTEF extraction and document codec, DOCX property rendering, and Markdown
formula rendering. The glossary registers the `formula-raw` profile and all
required parser/writer/conversion signatures. The API glossary records the
new policy and the cross-format boundary rule.

## Source checks

Before integration, the leaf was checked with direct `rustfmt --edition 2024`
over all changed Rust sources, `python3 -m json.tool tools/api_glossary.json`,
and tracked-source searches for every old zero-policy constructor and helper
call. Cargo/rustdoc/native qualification is intentionally left to the
coordinator's heavy-build slot; no benchmark artifact or binary is committed.

## Tests

Existing MTEF/OMML wire and semantic round-trip tests were migrated without
changing their expected records or AST values. New boundary tests cover exact
input admission, one-byte-over input refusal, caller-selected XML depth,
pre-emission OMML/MTEF output refusal, and writer work/item/string ceilings
with typed limit kinds.
