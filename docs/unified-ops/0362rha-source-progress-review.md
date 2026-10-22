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
