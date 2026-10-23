# 0362rha — retain source-failure progress during error classification

- Reviewed candidate: `273aa07748674cc67b6d6a229d9f8b7248e38762`.
- Decision: return the ODP freshness conversion for a fidelity repair; the
  remaining ODF classifications require their focused runtime checks.

The candidate correctly separates stale patch bases from source mutation.
Its `source_append::publication_error` helper, however, converts
`SourceContentPublicationError::SourceChanged` and nested source errors into
a bare `Error::SourceChanged`. The previous `Error::Plan` wrapper exposed the
original progress through `Error::progress`; the new variant returns `None`.
Classifying a freshness failure must preserve this public observation.

Retain the original progress when introducing a typed freshness variant, or
keep the original wrapper and classify its nested error at the public
boundary. Cover failures before and after progress is recorded, as well as
the distinction between `PatchConflict` and `SourceChanged`. Keep actual
source I/O failures separate from a wrong recorded base and preserve error
precedence when callbacks and source checks both fail.

The initial candidate was committed without Cargo validation because its
isolated checkout had no build cache. Focused tests and affected-consumer
checks are required before integration; the worker is performing them in a
separate bounded target. This review grants no testing or performance
exception.

## Follow-up review of `8efd67d9d0`

The repaired tail-append error retains progress and typed secondary callback
diagnostics. Boxing the diagnostic keeps the 104-byte error layout of its
pre-follow-up base; focused freshness tests and strict Clippy pass.

The boundary review still returns the candidate for two missing pieces.
`litchi-ops` has no source variant, `Classify` implementation or leaf conversion
for the public `litchi_odp::tail_append::Error`. Its new `SourceChanged`
therefore cannot yet deliver the advertised `io_error` through the facade.
Also, `SourceBackedPresentation::save` and deferred materialization convert
`SourceContentInsertionError` through `litchi_odf_common::Error`: freshness
loses progress, while `Replay` and `Transport` flatten typed diagnostics into
`Core::Other`. Direct `write_to` retains them, so the output routes differ.

Record 0362rhb must complete the facade plumbing and retain the original
typed diagnostics through save/materialization. Tests must obtain actual
public stale-source and stale-patch failures and check their final codes,
progress and secondary causes. Keep successful-path error layouts bounded,
migrate every affected consumer, and correct the earlier mapping claim.
