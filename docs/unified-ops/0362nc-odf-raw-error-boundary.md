# 0362nc — ODF raw XML error boundary

Status: source repair complete; native performance remains open.

`ResolvedReader::from_xml`, `read_resolved_event_into`, and
`read_resolved_event` now return `litchi_odf_common::Error`. Raw XML limit
failures are constructed at the reader boundary as
`Error::Limit { kind: LimitKind::RawXml(RawXmlLimitKind::...), ... }`, so
namespace scope depth, declaration count, active bindings, bytes, token
bytes, and the other dimensions retain their actual kind before any legacy
consumer converts them to a common-core error. ODT and ODS adapters perform
that conversion explicitly at their own boundary.

Validation:

* `cargo check -p litchi-odf-common -p litchi-odt -p litchi-ods --all-targets --all-features --offline`
* `cargo test -p litchi-odf-common --test resolved_reader --offline` — 15 passed.
* `cargo fmt --all`

The exact raw logs are retained outside the checkout under
`.coord/w8-odf-raw-xml/`: `clippy-final.log` (269 bytes, SHA-256
`30c3a30597f3e934f229a88109531a9fcf86ff73cbb746642fd22b017864a28a`) and
`libs-final.log` (171190 bytes, SHA-256
`f46ca4ddc89a33ba4288c136e437e1500c150b2ba0a65b834be9027aa7e6fc6b`).
