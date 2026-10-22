# 0362rm — DOCX raw codec boundaries

- Date: 2026-10-22.
- Status: source and tracked-consumer migration complete; runtime, allocation,
  and performance qualification remain pending the coordinator's released
  Cargo lane.
- Parent: 0362rc/0362r and the 0351 API contract.

The DOCX free package mutators and detached XML codecs now have one explicit
raw boundary. The normal package owner continues to use its focused `Edit`
families and publishes a package `Commit`. Direct callers that intentionally
work with an `OpcPackage` or detached codec use the corresponding `raw`
module and pass that family's typed `&Limits` policy.

## Public boundary

The following raw paths retain the historical capability with required typed
limits:

| Capability | Raw entry point |
| --- | --- |
| chart graph load/store | `chart::raw::{load,store}` |
| font XML and document font graph | `font::raw::{parse,write,read,put,remove}` |
| glossary XML and complete opaque graph | `glossary::raw::{read,write,load,put,remove}` |
| web-settings XML and package graph | `web::raw::{parse,read,write,put,remove}` |
| stylesWithEffects package graph | `styles::effects::raw::{load,put,remove}` |

Detached stylesWithEffects resources use `Resource::new(xml, &Limits)` and
`Snapshot::from_xml(owner, xml, &Limits)`. Web snapshots use
`web::Snapshot::from_xml(xml, &Limits)`. The former public
`Resource::from_xml` spelling and the old free package codec exports were
removed; all tracked consumers now use the canonical raw path or owner edit.
The public chart and glossary source snapshots also require their raw policy:
`chart::Snapshot::{load,read}(package, document, &Limits)` and
`glossary::Snapshot::{load,read}(package, &Limits)`. Their package reads pass
through the same raw admission checks before graph allocation.

The glossary graph remains transferable without flattening its opaque parts.
`glossary::raw::Part::new`, `replace_data`, and `set_relationships` also take
the required policy. Name, content-type, payload, relationship-count, and
relationship-metadata bounds are checked before the part retains new strings,
payload ownership, or a replacement relationship vector. Raw graph publication
checks newly serialized catalog XML once before staging it, and web/font raw
publication reuses the same serialized bytes for policy admission and package
staging. Untouched package payloads continue to be shared by the existing OPC
graph code.

The package modules retain only crate-private canonical implementation helpers;
there is no public alias for a former top-level `put`, `remove`, `store`,
`parse`, `read`, or `write` bridge. The raw XML scanners keep the format hard
ceilings while enforcing the caller's selected byte, depth, node, attribute,
event, text, graph, and metadata bounds.

## Static validation

- `rustfmt --edition 2024` passed for every changed Rust source file.
- `git diff --check` passed.
- Static searches found no tracked consumer of the removed public free codec
  exports or `Resource::from_xml`.
- Cargo, rustdoc, the API glossary capture, and native performance runs were
  intentionally not started while the coordinator's target and host-disk hold
  was active. The coordinator must run the full all-feature DOCX consumers,
  glossary/inventory capture, Clippy, and the paired raw lifecycle probes
  before acceptance.

No executable, fixture, or large performance artifact is part of this record.
