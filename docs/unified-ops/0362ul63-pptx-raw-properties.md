# 0362ul63 — PPTX raw properties and model3d policy propagation

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Base: `a942320908` (`uo2/coordinator-replay-integration`).

This slice gives the PPTX view-properties and presentation-properties codecs
one explicit, caller-selected limits contract. The public associated
`parse`, `parse_with_tolerance`, `parse_strict` and `to_xml` methods all take
their module-local `Limits`; the relationship loaders and the
source-preserving `read_only_recommended::Snapshot::{from_xml,from_xml_owned}`
take the same required policy. The package and slide owners derive those
policies from their retained `SourceLimits`, while direct raw callers choose a
policy explicitly. No default-only twin, deprecated alias or `_with_limits`
entry point was added.

The policy bounds input and output bytes, XML depth and nodes, decoded strings
and attributes, guides/slides, presentation extensions/colors, and model3d
XML, tree, relationship, link and embedded model/preview payload budgets.
Model3d package reads, selected-shape reads, edits and source relationship
validation carry the owner policy through every load, replacement and replay
path. Output admission is exact-counted before a fallible, exactly-reserved
writer copies retained opaque XML; source splices likewise reserve only their
checked final size. Unknown XML, MCE tolerant/strict behavior, namespace
mapping for generated package roots, relationship fidelity and existing
semantic validation remain intact.

Public required-limit signatures covered by this record:

- `litchi_pptx::view_properties::ViewProperties::{parse,parse_with_tolerance,parse_strict,to_xml}`;
- `litchi_pptx::view_properties::{load_from_package,load_from_package_with_tolerance}`;
- `litchi_pptx::presentation_properties::Properties::{parse,parse_with_tolerance,parse_strict,to_xml}`;
- `litchi_pptx::presentation_properties::{load_from_package,load_from_package_with_tolerance}`;
- `litchi_pptx::presentation_properties::read_only_recommended::Snapshot::{from_xml,from_xml_owned}`;
- `litchi_pptx::model3d::codec::{read,write}` and
  `litchi_pptx::model3d::validation::validate`.

The PPTX minimal and full glossary profiles register these signatures. Source
validation performed here is `rustfmt --check`, `git diff --check`, JSON
parsing of the glossary, and focused source/API sweeps. Cargo tests, rustdoc,
clippy and ABBA measurements are intentionally outstanding for the
coordinator's independent run; no binary or large performance evidence is
committed.
