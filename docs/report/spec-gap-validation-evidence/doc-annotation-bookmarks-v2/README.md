DOC annotation-bookmark facade and count closure
================================================

Adds Document::annotation_bookmarks(), returning a borrowed optional inert tag table. The selected source range is bounded and checked before allocation; strict parsing and its error are cached. Ranged comments also validate this table eagerly through the existing comment reader, which is stated in the accessor documentation.

MS-DOC section 2.9.277 limits annotation bookmarks to 0x3FFB despite the generic extended-STTB field admitting 0x3FFC. The annotation owner, comment reader and writer now share the structure-specific limit. The writer checks the count before allocating its ordered story metadata. Tests cover a complete maximum-size table, a correctly sized maximum-plus-one table, writer round-trip at the maximum, and writer rejection above it.

Independent corpus_binding_review found the v1 reader/writer inconsistency and insufficient boundary regression. V2 corrects both. Root reviewed the complete delta against the local specification and verified the new end-to-end boundary tests. Full final-source root gates passed 1,240 ordinary tests and 14 doctests (two ordinary tests and 12 doctests ignored), strict all-target/all-feature Clippy, warning-denied documentation and scoped formatting. This is crate-scoped evidence, not whole-workspace certification.
