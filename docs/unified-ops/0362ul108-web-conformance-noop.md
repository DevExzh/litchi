# 0362ul108 — Web Extensions conformance no-op boundary

The e130 runtime failure was in the shared Web Extensions planner. A semantic
no-op was rejected when an add-in omitted an unused relationship namespace
prefix or when vendor opaque text contained the other relationship dialect's
URI. `source_matches_conformance` treated every namespace declaration in a
retained fragment as a profile token, so the planner fell through to the
unsafe canonical-rewrite guard and returned `Invalid` instead of preserving
the package byte-for-byte.

Profile matching now examines only resolved relationship-valued attributes
(`r:id`, `r:embed`, and `r:link`). A part with no such attribute contributes no
profile signal; an actual relationship attribute must use the requested
namespace, and the package relationship image type remains checked separately.
Unused prefixes and vendor text therefore take the source-bound exact no-op
path, while a real Strict/Transitional relationship mismatch still forces the
existing changed/rewrite or refusal path. The shared planner is the only
production change; its DOCX facade consumer continues to use the same patch
and commit flow, and the existing DOCX conformance regression covers both
profiles and exact archive-member fidelity.

This is a source-only repair. No Cargo, build, rustdoc, runtime, native,
metadata, or performance command was run; the coordinator must rerun the
affected common and DOCX web-extension gates.
