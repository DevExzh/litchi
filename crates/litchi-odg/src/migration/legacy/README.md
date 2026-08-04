# ODG legacy migration

The files beside this document are the source-preserving relocation of the
former `litchi-odf::odg` implementation. They remain unwired because the
legacy parser and mutable editor import the umbrella's presentation parser,
shape vocabulary, and package extensions.

When migrated, parser code belongs in `codec`, drawing values in `model`,
package transactions in `package`, and edits/builders in `authoring`. The
deferred style-resource regression remains under `tests/deferred/legacy` and
is intentionally outside Cargo's direct integration-test discovery.
