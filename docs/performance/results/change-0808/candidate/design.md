# 0808 direct namespace-resolver candidate

This is an archive-only source candidate based on commit
`d28e3dc702d84f8752a2e329c2d3fc15f5c94f06`. It changes only the archived copy
of `crates/litchi-pptx/src/notes/codec.rs`; the production worktree remains
unchanged. The before leg is an exact copy of the production source at that
base. The after leg is not a production adoption or a performance claim.

The candidate replaces the production scanner's `NsReader<&[u8]>` event
transport with a direct `quick_xml::Reader<&[u8]>` and a local public
`NamespaceResolver`. It preserves the resolver state transitions that
`NsReader` performs: a pending scope is popped before every read; every Start
and Empty event is pushed before the existing depth, node, root, and element
inspection checks; Empty and End events set the pending pop before their
existing validation branches. A `NamespaceError` is converted through
`quick_xml::Error` before the shared `xml_error` adapter. The inspector receives
the resolver directly. The test-only counter helper keeps its existing
`NsReader` setup and passes its resolver to the adapted inspector.

The buffered scanner and `inspect_element_oracle` remain byte-for-byte
untouched. Existing attribute checking, namespace declaration handling,
relationship collection, XML limits, refusal ordering, and root classification
remain in place. The candidate adds focused differential coverage for nested
Start/End scopes, Empty scopes, default and prefixed rebindings, scope pop
restoration, reserved namespace errors, declaration-cap ordering, and a
reserved-namespace error at the depth boundary. All new cases compare the
direct scanner's values and refusal text with the retained buffered oracle.

The source review was guided by accepted ADRs 0001, 0002, 0005, 0006, 0008,
0011, 0013, 0024, 0030, 0031, and 0032, together with the normative ADR
index. Those constraints require correctness, bounded validation, preserved
typed refusals, crate ownership, and measured evidence before any adoption.
No new dependency, unsafe code, public API, ambient state, or workflow claim
is introduced by this archive.

Root owns the archived rustfmt pass, source/build checks, differential tests,
any native profiling or workflow capture, the independent review, production
restoration audit, and the final disposition. Root formatted the archived candidate and refreshed the manifest source and
production-path patch identities before application. Both buffered oracle
functions remain byte-identical after formatting.
