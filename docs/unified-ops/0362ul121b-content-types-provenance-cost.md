# 0362ul121b — Bound same-byte content-types provenance validation

Review of 121a found that `try_replace_content_types` called
`source_content_types_source()` unconditionally.  That could repeat the
package-part and relationship-override coverage walk even when the replacement
bytes were different, adding work and a possible allocation/limit failure to
ordinary changed publication.

The provenance lookup is now short-circuited unless the current and
replacement `[Content_Types].xml` bytes are identical and the package still
has a raw source XML allocation with those exact bytes.  Only that case needs
to distinguish a valid retained token from a missing or invalid token; the
latter still installs the checked replacement token.  Changed replacements,
missing/different raw source XML, and signed same-byte no-ops do not pay the
second coverage walk.  Existing validation, signature, and limit ordering is
otherwise unchanged.

This is a source-only follow-up to 121a.  No Cargo, build, rustdoc, runtime,
native, metadata, or performance command was run; the coordinator must rerun
the scoped Ribbon/common gate and its performance checks.
