# 0362na — ODF raw XML policy propagation

This continuation implements the owner-policy portion of the 0362n raw XML
review. It keeps the change required work; it does not add a conformance
exception.

## Raw reader and typed failures

`ResolvedReader::from_xml` now requires `XmlStreamLimits`. The reader applies
the retained policy while events are read, including input bytes, depth,
events, attributes, token bytes, namespace scope depth, declarations,
bindings, aggregate namespace bytes, prefix bytes, and URI bytes. Namespace
failures retain their raw dimension and observed/maximum values instead of
being inferred from an error string. The ODF common boundary reports
`LimitKind::RawXml(RawXmlLimitKind::...)`; legacy ODT catalog consumers still
receive the coarse core resource classification where that boundary is part
of their existing API.

The ODT fused reader and source-backed materialization retain the caller's
`XmlStreamLimits` through open, edit, commit, source-backed reopen and patch
replay. The historical sequential oracle continues to assert its old
`InvalidFormat` disposition, while the fused path asserts typed raw depth
observations. Namespace edge tests retain accept/reject parity with the old
reader and the raw reader tests assert typed namespace declaration and scope
dimensions.

`OpenOptions::default()` now has the same owner XML profile as
`OpenOptions::new()`. Direct raw reader `Default` remains separate so a raw
caller must choose its policy explicitly.

## ODS owner propagation

ODS `Archive` retains the XML policy used at ingress. Replacement packages,
worksheet snapshots, facade commits, definitions, DataPilot, database-range,
chart, and RDF snapshots and edits all carry that policy through internal
reopens and typed readback. Fresh byte constructors and authoring builders
continue to select the documented owner default; an opened owner never falls
back to that default during edit, commit or replay.

The regression `edited_and_replayed_owners_keep_explicit_xml_policy` opens an
ODS package with a 64 KiB token policy, edits it, commits it, replays the patch,
and verifies the same typed `ODF raw XML token bytes` refusal on the original,
edited and replayed owners.

## Validation and dispositions

- `cargo test -p litchi-odf-common --offline`: 363 library tests, 12 corpus
  tests, 15 splice tests, 2 relocation tests and doctests passed.
- `cargo test -p litchi-odt --lib --offline`: 807 passed.
- `cargo test -p litchi-ods --lib --offline`: 696 passed.
- `cargo fmt --all -- --check` and `git diff --check` passed.

The before/after ODT depth disposition is intentional: the old sequential
consumer remains `InvalidFormat: ODT content.xml nesting exceeds maximum depth
of 4096`; the new raw boundary is `ResourceLimit(Resource::Depth)` with
observed `4097`, maximum `4096`, and scope `ODF raw XML depth`. The raw common
boundary separately preserves `Depth`, `NamespaceDeclarations`, and
`NamespaceScopeDepth` kinds. No native performance qualification was run in
this continuation.
