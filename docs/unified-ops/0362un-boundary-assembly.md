# 0362un — boundary assembly repairs

The boundary gate’s Numbers identity audit was stricter than the helper it
inspected: `ArchiveObject::same_content_ignoring_offsets` compared retained
source provenance through `source_headers()`, so the helper body did not expose
the raw and canonical wire-header fields to the source-topology audit.  The
audit now follows that shared helper, requiring a comparison between the
distinct `self` and `other` operands and requiring `source_headers()` itself to
retain both accessors.  The production comparison remains allocation-free and
still ignores only physical offsets; no decoded-only comparison or provenance
loss is permitted.

Cargo metadata also reported `litchi-ops -> soapberry-zip` as a normal optional
workspace dependency.  The edge is used by production replay-publication
errors under the ODF features, so it is a canonical runtime edge rather than a
dev-only exception.  The checked-in policy now classifies it accordingly, and
the boundary tests cover both the normal/optional metadata classification and
the explicit raw and canonical provenance requirements, including rejection of
same-endpoint comparisons.

Validation was source-only: the focused boundary audit and the affected
`BoundaryPolicyTests` passed, as did rustfmt, JSON parsing, Python compilation,
and `git diff --check`.  The complete policy-test class was also attempted;
three unrelated checks require iWork native fixtures omitted by this sparse
source-only checkout.  No Cargo target or new target was started.
