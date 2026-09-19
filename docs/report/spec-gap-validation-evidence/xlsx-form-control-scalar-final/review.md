# Read-only review: XLSX form-control scalar lifecycle

Review date: 2026-09-19. This review covers the paired `x14:formControlPr`
and VML `ClientData` scalar lifecycle, the ordinary worksheet facade, source
publication, patch replay, conflict planning, limits, and the local evidence
boundary. It does not review the unrelated Custom Data implementation.

The implementation has a sound bounded shape. F-01, F-02, and F-03 below are
closed after the final source and focused-test fixes. The retained generated-
buffer resource audit is also closed against the frozen source and focused
tests. No independent production blocker remains in this review; the root
agent owns the separate workspace-wide gate receipt.

## Findings and closure

### F-01 — Source-only formula lifecycle (closed)

The mirror evidence explicitly treats the retained `#REF!` pair in
`tdf134769.xlsx` as preservation-only. It says that a new `fmlaLink` value must
meet the cell-reference rule and that source-error/presence changes are
read-only (`docs/report/spec-gap-validation-evidence/xlsx-form-control-mirror-evidence.md`,
the formula discussion and the `fmlaLink` matrix row).

The paired mirror checks `source_only` on both current values before a changed
`FmlaLink` replacement, and `validate_scalar` now admits a source-only value
only when it exactly equals the typed source projection. Detached/new
authoring remains subject to the leaf formula validator.

Required behavior is:

* preserve a source-only formula byte-for-byte during an unrelated scalar
  edit;
* permit an exact no-op for the retained source-only pair; and
* refuse a changed `FmlaLink` when either paired current value is source-only.

The focused test
`source_only_formula_allows_exact_noop_and_unrelated_scalar_but_refuses_replacement`
passes. It verifies exact no-op publication, preservation during an unrelated
scalar edit, refusal of a changed authored formula, and empty refusal output.
This closes the source-only formula finding.

### F-02 — Namespace-aware VML locator (closed)

`find_client_data` now uses `NsReader`, requires the VML namespace for direct
`shape`, the Excel namespace for direct `ClientData`, resolves an unqualified
`id`, decodes that value, and rejects duplicate target shapes/client-data
occurrences and malformed attributes. The focused tests cover a foreign
ClientData subtree, a foreign shape with a duplicate local id, and an XML
escaped shape id. The `source_scalar_` test group passes 11/11, closing the
wrong-range finding.

### F-03 — Focused integration compile gate (closed)

The focused command run during this review was:

```text
cargo test -p litchi-xlsx --test form_control_scalar_lifecycle --no-fail-fast
```

The production library passes `cargo check -p litchi-xlsx --lib`; the
test-only mirror accessors are correctly gated; and the stale test calls to
`Commit::changed()` have been corrected. The full lifecycle target builds and
runs 32 tests, all of which pass, including the managed-budget retention and
source-only formula cases.

The changed incoming-read-set test now passes, including its no-mutation
assertion.

### Resource ownership audit — closed

The frozen implementation has a bounded ownership path for both source-backed
and ordinary scalar results:

* A changed or no-op mirror pair transfers one shared lease covering output
  bytes, the two retained `Arc`/`Vec` headers, the hold itself, and the object
  charges. Forward patches, inverse patches, result clones, and opaque aliases
  keep that lease alive; managed payloads cannot escape through an uncharged
  shared `Arc` path.
* Generated `Properties` retain their semantic model lease and their
  `SourcePayload`; parser aliases are accounted once by source identity.
  Generated collections carry the projection lease through collection clones,
  including URI/string storage. Managed before-payloads continue to share
  their existing source handles.
* Failed owner admission does not retain the deferred package read set. The
  package read set is attached only after owner admission, while caller
  content-types limits preserve the typed owner-limit refusal. This keeps
  mirror-node, MCE, and opaque-cap failures bounded and correctly ordered.
* Eager materialization uses bounded copies, while source publication and
  replay retain exact source authority, before-bytes, and owner read-set
  checks. The ordinary `Patch::apply_to` path itself does not rescan the form
  owner; the complete owner readback occurs during initial ordinary commit.

The focused lifecycle run is `32 passed, 0 failed`, including
`managed_scalar_result_handles_retain_and_refund_memory_and_objects`, and the
focused owner-read run is `39 passed, 0 failed`. These close the resource
ownership finding for this profile.

## Accepted design and scope

The exact immutable `PatchAuthority` restriction is appropriate for the first
ordinary form-control profile. The forward patch is pinned to the source
`Workbook` allocation and its inverse to the corresponding result; a fresh
equal-byte workbook cannot satisfy the authority. `Patch::apply_to` checks
unsigned/protected policy and exact physical before bytes, then rebuilds the
workbook. The initial ordinary transaction performs the complete form-control
owner readback before returning its patch; subsequent ordinary replay relies
on the exact source authority plus before-bytes and does not independently
rescan the form-control owner. The durable form retains its complete
serialized-source precondition. This is stricter than a payload-only patch
and matches ADR 0003's immutable snapshot/read-set model.

The source-backed path retains the full owner read set. Replay compares the
worksheet, DrawingML, VML, selected properties, sidecar targets, root
relationships, content types, signatures, incoming edges, source version, and
MCE provenance, excluding only the two payload members that the operation is
allowed to replace. Signature refusal occurs before source publication, and
the candidate is re-read through the complete owner before a stream receives
bytes.

The ordinary facade has the intended bounded scope: one control per worksheet
batch, at most the owner scalar-operation limit, two changed physical members,
paired output/staging limits, and atomic candidate validation. `PackageChange::FormControl`
records sheet/control identity and before/after properties, and its inverse
swaps those states. Same-sheet independent form-control branches produce
`Conflict::FormControls`; the three-way planner projects the same
`sheet/{position}/form-controls` write key. Cell edits remain composable. The
DrawingML transfer path is separately guarded by its existing drawing-owner
preconditions, and a transfer target with an existing drawing owner is refused,
so no additional form-control/drawing join rule is required for this profile.

The narrowed ordinary publication fallback is also correct in scope:
`try_replace_owned_xml_part_bytes` is used only for parts whose content type is
`CONTROL_PROPERTIES_CONTENT_TYPE`. It checks the expected bytes, validates the
replacement with the current XML content type and limits, and installs source
provenance before the paired VML member is staged. Keeping the generic fallback
out of unrelated XML avoids coupling other worksheet edits to this profile.

Mirror lexical/default behavior matches the local evidence where the profile
admits a row: VML booleans use only `t`, `f`, `true`, `false`, empty, `True`,
and `False`; object-type bridging stays at the three fixture-backed pairs;
`checked` and `editVal` use field-specific mappings; and `lockText` treats VML
absence as effective true rather than x14 false. Effective omitted fields are
not used to invent authored VML fields. `EditVal` has specification proof but
still needs a retained edit-control fixture before a writer claim. Formula
graphs, `FmlaGroup` relocation, list items, `FmlaRange`, and control
creation/deletion remain read-only/out of scope.

## Dependency and audit scope

The form-control path depends on generic OPC source XML/topology and read-set
APIs, including the control-properties source-proof helper needed for ordinary
replay. The reviewed changes contain no Custom Data imports or behavior and do
not require the unrelated Custom Data owner to be complete. The final isolated
gate should include the compatible generic OPC source-publication changes;
Custom Data failures in a workspace-wide run should remain a separate pending
feature result.

`docs/report/spec-gap-audit.md` §5 still describes `formControlPr` as zero
matches with only a generic metadata row. Once F-01/F-02 and the gates pass,
update that row to describe this bounded inert scalar lifecycle and retain the
explicit exclusions above. Do not mark control execution, rendering, graph
formula/list editing, creation/deletion, ActiveX behavior, or native Office
acceptance as implemented.

## Verification status

The preliminary OPC review in `opc-review.md` reports passing focused source
content-types, signature, relationship, limits, cancellation, and crate-check
tests. The current `litchi-xlsx` library check passes and all 1,184 library
tests pass. The exact formula and namespace-focused tests pass, the full
lifecycle target passes 32/32, and the owner-read target passes 39/39,
including the finite-budget retention/refund regressions. The selected source
hashes match the final freeze receipt. The root agent is completing the
workspace-wide gate receipt separately; no form-control semantic or resource
blocker remains in this review.
