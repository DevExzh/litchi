# 0362ne — Require ODF style and master XML policies

- Date: 2026-10-22.
- Parent: 0362n and the typed raw-error continuations through 0362nd.
- Status: implemented source slice; native performance qualification remains
  open.

The shared ODF data-style and master-page XML APIs now require an
`XmlStreamLimits` value. Their readers use the retained `ResolvedReader`, so
input bytes, depth, events, attributes, token bytes, namespace declarations,
scope and URI dimensions are checked by the caller's policy. The old implicit
XML-byte, depth and event ceilings were removed from these parser paths;
semantic style and master-page bounds remain in force. The public style-data
operations and master-page writers validate and scan with the same policy
before publishing a changed XML string.

Every affected consumer was migrated in this slice. ODT's header/footer
snapshot retains the policy used at parse time and passes it through master
page reads, edits, commits and patch replay for package and flat owners. ODT
document projections and mutable semantic style operations use that retained
policy. ODP's public `parse_masters` accepts the required policy; package
authoring and validation use the documented owner default because they create
fresh owner state. Corpus, package, flat, transaction and replay callers were
updated in the same source change.

The common data-style and master reader tests include boundary checks that
assert the typed `RawXmlLimitKind::InputBytes` and `RawXmlLimitKind::Depth`
observations. Existing ODF, ODT and ODP fidelity and mutation tests continue
to exercise exact XML preservation, schema validation, transfer, commit and
replay behavior. The remaining 0362n surfaces (metadata, manifest, drawing
resources and other raw XML families) stay tracked work; this record does not
turn them into conformance exceptions.

Validation:

* `cargo fmt --all` passed.
* `cargo check --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-targets --all-features` passed.
* `cargo test --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-features --all-targets` passed.
* `cargo clippy --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-targets --all-features -- -D warnings` passed.
* `git diff --check` passed.

No native performance measurement or acceptance claim was made in this slice.
