# 0362ul62b — PPTX raw-content replay correction

Status: source-ready correction to `0362ul62a`, based on the coordinator
replay integration. Native Cargo validation, the workspace gate, and ABBA
remain coordinator work because this leaf had no build slot.

## Corrections

- Source-backed text-box staging now passes the selected shape policy through
  the five-argument composition seam.
- Notes replay validates the recipient aggregate ceiling over the retained
  notes master, referenced theme, and notes slides after changed XML members
  pass their individual typed checks. The aggregate inventory is byte-only and
  does not reparse the owner graph.
- Media replay tracks changed `/ppt/media/` resources and revalidates every
  unchanged slide that references one, so `max_total_payload_bytes` applies
  to resource-only cross-copy changes as well as changed slide XML.
- Comment-author presence and comment-threading replay now run the typed
  collaboration locator under the recipient comment limits in addition to the
  generic legacy-comment grammar validator.
- Added exact/one-under aggregate coverage for notes and media resource-only
  replay, plus typed presence and threading extension validation tests.

The shared `Arc<SemanticLimitsRef>` carrier, common source rebase/admission,
and generic `PartPatch` helper remain coordinator merge responsibilities. No
generated evidence, fixture, executable, or `3rdparty/` path is included.
