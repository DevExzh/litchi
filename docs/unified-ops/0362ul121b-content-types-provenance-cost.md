# 0362ul121b — Bound same-byte content-types provenance validation (returned; see 121c)

Review of 121a found that `try_replace_content_types` called
`source_content_types_source()` unconditionally.  That could repeat the
package-part and relationship-override coverage walk even when the replacement
bytes were different, adding work and a possible allocation/limit failure to
ordinary changed publication.

The 121b provenance lookup was short-circuited unless the current and
replacement `[Content_Types].xml` bytes are identical and the package still
has a raw source XML allocation with those exact bytes.  Only that case needs
to distinguish a valid retained token from a missing or invalid token; the
latter still installs the checked replacement token.  Changed replacements,
missing/different raw source XML avoid the second coverage walk.  Review then
found that signed state was computed after this gated lookup, so an exact
same-byte signed no-op could still scan.  Follow-up 121c moves the signed
no-op return before provenance validation while preserving changed signed
refusal ordering.

This is a source-only follow-up to 121a.  No Cargo, build, rustdoc, runtime,
native, metadata, or performance command was run; the coordinator must rerun
the scoped Ribbon/common gate and its performance checks.
