# 0362rk — PPTX and OOXML metadata error taxonomy

This repair completes the PPTX and shared OOXML consumers of the 0351 error
contract. The original review commit was `d7ba6f3c98`; this commit amends it
with the perf-baseline consumers found during the integration audit.

The common OOXML property transaction now returns unit `PatchConflict` for a
caller-supplied or stale property base, patch composition mismatch, or scope
mismatch. The two post-validation disappearance cases return
`Internal(&'static str)`, preserving the distinction between stale input and
an implementation invariant. PPTX metadata consumers and the ops classifier
map `PatchConflict` to `invalid_input`.

PPTX slide-copy, slide-removal, slide-order, shape-transfer, and duplicate
selector routes now return `Refused(Refusal)`. Their retained reason enums
provide the shared `RefusalKind`, and each route supplies a fixed,
content-free reason. Recognized superseded durable revision headers also
return unit `PatchConflict`. Ambiguous slide bindings remain `Invalid`, so
malformed source data keeps the `corrupt_file` classification. The perf
baseline's cross-copy and slide-boundary checks use the same refusal shape.

The removed typed plan errors carried an enum and an owned `String`; the
replacement refusal carries the shared static `Refusal` fields, and replay
conflicts are unit variants. No success path adds an allocation. Metadata
replay tests cover PPTX, XLSX, XLSB, and DOCX, while the ops tests cover
stale-base classification and the PPTX special classifications.

Validation was source-only under the coordinator's disk hold: targeted
`rustfmt --check` for the changed crate files, boundary file, and shape test,
`git diff --check`, and `git grep` across the full tracked tree passed. The
perf-baseline library retains one pre-existing formatting difference outside
these hunks. The repository size and binary guards passed. Cargo was not run
and no compiler cache was created; the coordinator must run the warmed
combined gate.
