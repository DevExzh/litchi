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
* `cargo test -p litchi-odf-common -p litchi-odt -p litchi-ods --all-features --offline`
  — 363 common unit tests, 696 ODS unit tests, 807 ODT unit tests, all integration
  tests and doctests passed (one common doctest is intentionally ignored).
* `cargo clippy -p litchi-odf-common -p litchi-odt -p litchi-ods --all-targets
  --all-features --offline -- -D warnings` passed.

The exact raw logs are retained outside the checkout under
`(deleted-coord-artifact)`: `clippy-final.log` (269 bytes, SHA-256
`30c3a30597f3e934f229a88109531a9fcf86ff73cbb746642fd22b017864a28a`) and
`libs-final.log` (171190 bytes, SHA-256
`f46ca4ddc89a33ba4288c136e437e1500c150b2ba0a65b834be9027aa7e6fc6b`).
The revised validation logs are `clippy-final-revised.log` (387 bytes, SHA-256
`4bddd61e4c85215bf5ebe691a7455c44d959fede9c0ef4ef341966144469e91f`) and
`libs-final-revised.log` (329625 bytes, SHA-256
`ac5b06456e331c35192f0f58aaed1c1fa18cb2434361e1190bddadc7e11e8608`).
