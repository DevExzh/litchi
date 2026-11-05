# 0362ul62a — PPTX raw-content replay and source-policy review

Status: source-ready follow-up to `0362ul62`, based on coordinator replay
integration tip `a942320908`. Native Cargo validation, the workspace gate and
ABBA remain coordinator work because this worker had no build slot.

## Review fixes

- Source-backed presentation editors, primary projections, prepared topology
  snapshots, focused edits, and test snapshots now retain the complete
  ingress semantic bundle. A snapshot's direct transition projection uses the
  selected transition policy instead of silently restoring defaults. The
  profile is retained for destination admission; it is not part of exact
  source identity, so the shared source patch rebase can validate and rebind
  changed payloads under the recipient profile.
- The shared part-patch engine now offers changed-payload validation without a
  package-wide scan. Media, slide tags, shape tags, notes, classic comments,
  and classic-comment collaboration replays validate changed XML/payloads
  with the recipient package's selected profile before publication.
- The remaining bounded reservations in comment authors/comments, comment
  storage, media slide XML, notes close-tag scanning, notes relationship
  validation, and notes slide-source inventory are fallible and report the
  typed `Error::Allocation` variant.

## Coverage and integration notes

The existing exact/one-under raw-limit and fidelity tests remain in place; the
source regression `selected_semantic_bundle_survives_source_editor_and_snapshot_projection`
checks that an overridden content policy survives both source-owner
projection stages. The replay validator visits only payload deltas and
therefore preserves the patch engine's relationship-only fast path. The
source fields use the current
content-only `SemanticLimits` carrier on this leaf; the coordinator must merge
them into the shared `Arc<SemanticLimitsRef>` bundle from the shape/motion and
properties leaves without overwriting sibling policies.

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
