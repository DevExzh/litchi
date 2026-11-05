# 0362ul121c — Order signed content-types no-ops before provenance lookup

The 121b gate still computed the retained-token match before checking signed
state.  A same-byte signed no-op with an exact raw source manifest could
therefore pay the content-types coverage walk and expose its allocation or
limit failure, despite requiring no publication.

`try_replace_content_types` now computes the signed disposition immediately
after the existing current-manifest and replacement-owner checks, returns for
a same-byte signed no-op, and only then performs the exact-raw-token lookup
needed by unsigned same-byte replay.  Changed signed replacements still reach
`SignedSourceRequiresExplicitPolicy`; unsigned missing-token destinations
still install the validated retained token.

This is a source-only follow-up to 121b.  No Cargo, build, rustdoc, runtime,
native, metadata, or performance command was run; the coordinator must rerun
the scoped Ribbon/common gate and performance checks.
