# 0362ul62f — PPTX raw replay index closure and deferred ingress

This return closes the remaining notes/media replay-index review findings for
the PPTX raw-content family.

The notes closure now retains every internal, non-external relationship whose
relationship type ends in `/theme`, including opaque extra edges and repeated
edges. Their byte contributions and reverse users are updated with the exact
changed-part set. A changed or newly admitted edge whose internal theme target
is absent is refused atomically, preserving the prior `PartNotFound` behavior;
the replay aggregate validator reports the same error for a malformed package
admitted before replay.

The retained maps use immutable path-copying AVL trees. Each changed key copies
only a logarithmic path, so replay lookup/update cost is independent of
publication history and no periodic full-index flatten retains old package
indexes. The initial index reads `PartMetadata::payload_len()` and relationship
metadata, rather than forcing every deferred part payload. A package-open
regression confirms that an unrelated slide remains undecoded. Media replay
also deduplicates slides reached through multiple changed resources before
loading each affected slide.

Focused source regressions cover all-theme-edge accounting, relationship-only
missing-theme refusal, authored index publication, resource-only media replay,
and deferred slide ingress. No public API was added or renamed in this return.

Validation for this source-only return is limited to `rustfmt --edition 2024`
on changed Rust files and `git diff --check`; Cargo builds/tests remain for the
coordinator's composed validation slot.
