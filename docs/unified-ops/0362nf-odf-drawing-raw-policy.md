# 0362nf — Require ODF drawing-resource XML policies

- Date: 2026-10-22.
- Parent: 0362ne and the typed raw-error continuations through 0362ne.
- Status: implemented source slice; native performance qualification remains
  open.

The shared ODF drawing-resource parsers now require an `XmlStreamLimits` value
for gradients, hatches, markers, stroke dashes, opacities and fill images.
Each parser uses the retained `ResolvedReader`, so raw input bytes and the
other XML stream dimensions are checked by the caller's policy. The previous
fixed input-byte ceiling was removed while the existing semantic resource
validation remains in force.

Every affected consumer was migrated in this slice. ODT package, flat and
mutable semantic style paths pass the owner's retained policy; ODP's public
drawing-resource projections use the documented owner default because they
create fresh projection state. Common parser tests assert the typed
`RawXmlLimitKind::InputBytes` result at the boundary, and ODT fixture tests
continue to cover exact resource fidelity for all six resource families.

The remaining 0362n raw XML surfaces stay tracked work; this record does not
turn them into conformance exceptions.

Validation:

* `cargo fmt --all -- --check` passed.
* `cargo check --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-targets --all-features` passed.
* The focused common boundary test passed:
  `cargo test --offline -p litchi-odf-common --test parser_hardening drawing_resource_parsers_report_the_required_raw_xml_policy`.
* The ODT drawing-resource fidelity tests passed:
  `cargo test --offline -p litchi-odt --test drawing_style_resource_fixtures`.
* `cargo test --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-features --all-targets` passed.
* `cargo clippy --offline -p litchi-odf-common -p litchi-odt -p litchi-odp --all-targets --all-features -- -D warnings` passed.
* `git diff --check` passed.

No native performance measurement or acceptance claim was made in this slice.
