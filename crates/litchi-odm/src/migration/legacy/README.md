# ODM legacy migration

The files beside this document are the source-preserving relocation of the
former `litchi-odf::odm` implementation. They remain unwired because the
legacy builder/editor imports ODT document types and umbrella-only section,
signing, and package APIs.

The next step must move those dependencies behind a neutral common/text layer,
then place parsing in `codec`, master semantics in `model`, package ownership
in `package`, and edits in `authoring`. The deferred ODM regression remains
under `tests/deferred/legacy` and is intentionally not a direct Cargo test.
