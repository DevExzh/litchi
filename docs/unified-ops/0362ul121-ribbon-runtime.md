# 0362ul121 — Ribbon replay runtime repairs (returned; superseded by 121a)

The a8 frozen runtime exposed two Ribbon publication defects.  A stale patch
could accept a recipient after an unrelated part was added because the source
check captured only the selected Ribbon dependency closure.  The initial
repair added a recorded package part-count guard, which made that fixture
conflict but was returned because unrelated package parts are valid replay
recipients under the selected-closure contract.  Follow-up record 121a removes
that over-broad guard and changes the regression to mutate the selected Ribbon
part for the true `PatchConflict` boundary.

The inverse of a removal that restores an unshared binary image also reached
the correct binary and XML parts, but its candidate was rejected when the
checked OPC content-types preservation API re-admitted an authored destination
manifest as a retained token.  Destination validation now accepts that
provenance transition only when the retained bytes are exactly the recorded
content-types bytes; an explicitly retained destination token still requires
the exact token.  This mirrors the existing relationship-token proof and
keeps the binary payload and opaque relationship closure checks unchanged.

The initial source-only repair was made against the frozen a8 capture.  No
Cargo, build, rustdoc, runtime, native, metadata, or performance command was
run; the coordinator must rerun the scoped Ribbon/common and dependent web
gates after 121a.
