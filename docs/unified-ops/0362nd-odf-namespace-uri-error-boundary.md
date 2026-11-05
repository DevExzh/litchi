# 0362nd — Carry ODF namespace URI limits as typed raw errors

Status: source repair complete; native performance remains open.

The namespace URI decoder now carries its own typed limit result from the
point where the URI byte budget is checked. `NamespaceUriBytes` is constructed
as `LimitKind::RawXml(RawXmlLimitKind::NamespaceUriBytes)` without creating a
core `ResourceLimit` or recovering a dimension from a scope string. The public
`namespace::normalize_namespace_uri` returns the common ODF `Error` directly.
Only the legacy `resolved_namespace_uri` projection converts the private
namespace error back to the established core error shape. Invalid limits,
lexical byte overflow and normalized URI overflow retain their exact
observation and maximum values; ordinary XML and allocation errors remain
wrapped as the ODF core error.

The public namespace tests accept an exact URI byte ceiling, reject one byte
below it with the typed `NamespaceUriBytes` dimension and check an invalid zero
ceiling at the same boundary. Existing escaped-URI, borrowed fast-path,
source-event, retained-reader-policy and default-envelope tests remain
unchanged and pass. No owner policy or namespace normalization behavior was
relaxed.

Validation:

* `cargo fmt --all -- --check` passed.
* `cargo test -p litchi-odf-common --test namespace_semantics --test resolved_reader --offline` — 8 namespace tests and 15 resolved-reader tests passed.
* `cargo test -p litchi-odf-common -p litchi-odt -p litchi-ods --all-features --offline` — 363 common unit tests, 696 ODS unit tests, 807 ODT unit tests, all integration tests and doctests passed (one common doctest is intentionally ignored).
* `cargo clippy -p litchi-odf-common -p litchi-odt -p litchi-ods --all-targets --all-features --offline -- -D warnings` passed.

The exact logs are retained outside the checkout under
`.coord/w8-odf-raw-xml/`: `namespace-uri-libs-final.log` (329690 bytes,
SHA-256 `a739fc424a17ff0a506b3852f81df98152596a1ecff7fb2afa43519a5c20c8ac`)
and `namespace-uri-clippy-final.log` (387 bytes, SHA-256
`b16356d59ffb7170ad0adb296f9a058457ee5fb0834ba1cb4d33b1895dc998c1`).
