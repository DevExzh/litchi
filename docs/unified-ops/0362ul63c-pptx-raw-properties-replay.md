# 0362ul63c — PPTX raw-properties replay policy

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Parent: `0362ul63b` (`fix(pptx): close raw-properties MCE and opaque-fragment review`).

This correction closes the replay-policy gap for the typed PPTX raw-property
families. `apply_math_properties`, `apply_model3d`, and every presentation
metadata `apply_*` family now use the shared recipient-aware replay hook.
Before publication, the hook validates only changed payloads under the
recipient package's selected `SourceLimits` (including an explicit
`OpenOptions::source_limits` on borrowed owning ingress): presentation-properties payloads
are strict-parsed, changed model3d slide XML is loaded with the recipient's
structural and payload profile while model and preview parts receive their
selected byte bounds, and metadata owners use their detached/package-aware
codecs after the physical recipient scan. A no-op still returns a duplicate
without invoking a semantic scan.

The read-only-recommendation owner edit already re-parses its candidate through
`load_snapshot`, which derives the same presentation-properties policy from the
recipient's selected source limits. Its public package replay now adds the
same strict changed-payload admission to the shared part-patch path. View-
properties is a raw projection with no patch family, so its required-limit
parse and serialization paths remain unchanged. The new boundary tests cover a
tighter recipient presentation profile and a model payload above the recipient
model budget; existing expansion-boundary and opaque-fidelity tests remain in
place. The
metadata replay hook covers Changes Information, color maps, custom shows,
Designer tags, show events, guides, revision information, slide
synchronization, presentation structure, and WebVTT tracks; Designer tags
retain their selected-owner transaction check in addition to the changed XML
physical-budget admission.

Source checks performed here are rustfmt and `git diff --check`. The
coordinator's approved F5 validation independently passed the focused PPTX
MCE, properties/model3d, all-target, ops and facade rows at source
`a75cf6419d`; the replay correction itself still requires the coordinator's
reviewed Cargo rerun. No binary or large performance evidence is committed.
