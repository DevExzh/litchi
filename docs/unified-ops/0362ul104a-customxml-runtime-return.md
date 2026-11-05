# 0362ul104a — Custom XML runtime refusal classification return

The frozen joint runtime source `ea5c95e212e1f8e934a0485eb0b1f51355f13ba5`
passed 526 of the 531 `litchi-ooxml-common` library tests. The one Custom XML
failure was `custom_xml::tests::failed_add_is_transactional`: an item and its
properties part used the same destination, and the publication boundary
returned a generic `Invalid` error while the established producer contract
requires a typed `RefusalKind::UnsupportedTarget` refusal.

`custom_xml::package::reject_part_aliases` now uses the existing
`unsupported_target` constructor for both source/data and properties/owner
aliases. The check remains before any package mutation, keeps the operation's
content-free reason, and leaves the caller package unchanged. No production
graph or payload behavior changes beyond the error classification.

This is a source-only repair. `git diff --check` and rustfmt for the changed
file pass. No Cargo, metadata, test, rustdoc, native, or benchmark command was
run on this leaf; the coordinator must rerun the bounded common runtime packet
against this exact commit. The four unrelated VBA failures remain with their
separate return leaf.
