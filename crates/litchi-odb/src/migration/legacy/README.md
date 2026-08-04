# ODB legacy migration

The files beside this document are the source-preserving relocation of the
former `litchi-odf::odb` implementation. They are intentionally not declared
by the active crate module tree yet: they depend on the former umbrella
package/document types and therefore cannot be wired without reintroducing a
family-to-umbrella edge.

The next migration step maps their XML codecs into `codec`, semantic values
into `model`, package ownership into `package`, and edits into `authoring`.
The deferred ODB integration tests live under `tests/deferred/legacy` and are
kept outside Cargo's direct test discovery until that mapping is complete.
