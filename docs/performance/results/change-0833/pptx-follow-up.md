# 0833 — PPTX opened-capture follow-up candidate

This is a read-only source and evidence review at HEAD `eefaca16e3`
(`perf(xlsx): keep first column assignment inline`). It records one bounded PPTX
candidate for consideration after the current filesystem-cold qualification
batch. The candidate remains unmeasured pending differential guards and matched
measurements; no speedup is claimed. The three unrelated workspace files
present during review remain outside this record.

## Candidate: validate the notes graph without materializing a discarded snapshot

`capture_internal` needs the notes loader's validation result, but it does not
retain the resulting notes snapshot. At
[`opened/model.rs:818-826`](../../../../crates/litchi-pptx/src/opened/model.rs#L818),
the result of either
`notes::load_snapshot_with_slide_root_proofs` or `notes::load_snapshot` is bound
to `_notes` and then dropped. The call path is:

1. `opened::model::capture_internal` resolves the presentation and slide
   references at lines 738-749, performs the slide identity, relationship and
   name checks at lines 767-817, and reaches the notes stage at lines 818-826.
2. `notes::load_snapshot_with_slide_root_proofs` at
   [`notes/package.rs:317-329`](../../../../crates/litchi-pptx/src/notes/package.rs#L317)
   runs `load_index_with_slide_root_proofs` and then calls
   `snapshot_from_index`.
3. `load_index_with_slide_root_proofs` at
   [`notes/package.rs:426-697`](../../../../crates/litchi-pptx/src/notes/package.rs#L426)
   performs the notes graph validation: presentation and slide conformance,
   slide inventory and relationships, notes master and theme ownership and
   content types, notes-slide backlinks, opaque relationships, aggregate byte
   limits and orphan checks.
4. `snapshot_from_index` at
   [`notes/package.rs:331-364`](../../../../crates/litchi-pptx/src/notes/package.rs#L331)
   then calls `materialize` at line 339. `materialize` at
   [`notes/package.rs:715-764`](../../../../crates/litchi-pptx/src/notes/package.rs#L715)
   copies the master, theme and notes-slide payloads through `own_blob`, builds
   a lifetime-free `Graph`, and collects relationship metadata.
5. The same function builds `PartState` values and invokes
   `Snapshot::from_parts` at lines 340-363. That sorts the parts, runs
   `validation::validate_parts` and computes the notes `fingerprint` at
   [`notes/transaction.rs:68-89`](../../../../crates/litchi-pptx/src/notes/transaction.rs#L68)
   and [`notes/transaction.rs:471-491`](../../../../crates/litchi-pptx/src/notes/transaction.rs#L471).
   The snapshot and its graph are then dropped by `capture_internal`.

The concrete follow-up is to split this use from the public/editable notes
snapshot path. A private opened-capture helper could consume the validated
`GraphIndex`, construct the same Arc-backed `PartState` metadata, run the same
`validate_parts` checks and compute the same notes fingerprint, then discard
those temporary states without calling `materialize` or `own_blob`. The public
`load_snapshot` and `load_snapshot_with_slide_root_proofs` paths would continue
to materialize the lifetime-free editable graph. This targets temporary owned
payload copies and graph construction; it is not a graph cache or a new
retained memo.

The helper must preserve the source-part identity check in
`Snapshot::from_parts`, the sorted relationship metadata and every limit/error
from `load_index_with_slide_root_proofs`. The notes fingerprint should remain
computed until a separate source review proves that the internal value is
observationally dead in opened capture. This record does not propose removing
any fingerprint work. In all cases, the complete package revision and part
digests in `opened/model.rs:827-851` must remain exactly on the capture path;
`package_fingerprint_with_memo`, signed-package handling, and its ownership
rules are mandatory.

## Evidence and attribution

The latest profile,
[`0829-pptx-edit-phase-profile.md`](../../0829-pptx-edit-phase-profile.md),
reports 23 reports and 4,549 verified samples (lines 3-7). Its current fixture
and workflow are stated at lines 16-29: `test-data/ooxml/pptx/shapes.pptx`, a
fresh package per iteration, public shape `(0, 0)` text edit, commit and apply,
with serialization, hashing and readback outside the edit timer.

The exact-owner phase table at lines 77-102 has transaction-capture samples
`1,196/1,193`, with exact edit-owner populations `2,670/2,685`. The selected
self-leaf table at lines 104-120 reports capture
`litchi_pptx::notes::*` as `103/101` samples. The same profile explicitly says
at lines 122-127 that these are not proof of inflated or copied bytes, and at
lines 136-155 that complete fingerprinting, limits, preservation and
publication validation remain mandatory. The profile also states that it
adopts no candidate and that the sample counts are neither wall-clock
fractions nor an Amdahl model.

The older real-deck record gives mechanism context, not current attribution:
[`0649-pptx-opened-transaction-real-deck-edit.md:174-195`](../../0649-pptx-opened-transaction-real-deck-edit.md#L174)
measured 15 notes/index/root/scan passes over 273,892 bytes and 38.95%
callgrind-inclusive `Ir` for the notes path. That record also says the
duplicate `slide_references` pass was only 2,357 of 819,319 MCE input bytes,
0.29%, about 0.18 ms of 128 ms (`0649:312-317`). The duplicate-catalog idea is
therefore a lower-value alternative and is not the candidate selected here.

The accepted 0693 notes-proof record confirms the existing boundary:
[`0693-pptx-capture-notes-proof.md:131-150`](../../0693-pptx-capture-notes-proof.md#L131)
retains presentation inventory, relationships, content types, conformance,
master/theme, notes-resource, orphan and snapshot-materialization checks. It
also says only the raw witness and conformance survive the capture-local proof.
This makes a blind deletion of `materialize` unsafe; the proposed helper must
retain equivalent source-part checks before dropping temporary state.

The benefit is therefore an **unmeasured hypothesis**. The 0829 notes self-leaf
counts show that notes work is present in capture, while the source call path
shows that a discarded snapshot performs extra copying and graph construction.
Neither establishes wall-clock, allocation, copied-byte or peak-memory savings
at `eefaca16e3`. Fresh attribution is required before implementation is
accepted: matched cold capture and commit phase measurements, allocation calls
and requested bytes, copied bytes, and peak retained bytes on the checked-in
fixture plus notes-bearing controls.

## Required preservation and guards

The candidate must keep the notes call at the current stage, after slide
identity/name checks and before `SlideNameIndex::build` at
`opened/model.rs:827`. It must preserve the following invariants.

- `load_index_with_slide_root_proofs` remains the source of truth for semantic
  refusal, including exact source-identity fallback for reordered equal-length
  slides. No refusal is memoized or inferred from a graph shortcut.
- Presentation, slide, notes-master, theme and notes-slide content-type and
  relationship checks remain complete. Opaque relationships, backlinks,
  one-to-one notes ownership, orphan checks, node/depth/attribute limits and
  the distinct notes-root/package byte ceilings remain active.
- `validation::validate_parts`, source identity and notes fingerprinting retain
  their current behavior. `PartState::from_part` may share the package's
  payload `Arc`, but the helper must not retain a parsed graph or create
  variable-sized derived state in the opened snapshot.
- `package_fingerprint_with_memo`, `Revision`, `PartDigests`, retained MCE,
  `SlideRootMemo`, `Patch::capture`, signature handling and final publication
  source checks remain unchanged. The staged commit recapture at
  [`opened/transaction.rs:1241-1297`](../../../../crates/litchi-pptx/src/opened/transaction.rs#L1241)
  must still fingerprint the complete staged package and capture the final
  snapshot before returning.
- Public notes loading and notes transactions continue to receive a complete,
  lifetime-free `Graph` with copied payloads. The optimization applies only to
  the opened capture validation use whose result is currently discarded.

## Differential test and measurement plan

Existing opened tests provide the required refusal and ordering oracle. Keep
and extend the independent legacy comparison at
[`opened/tests.rs:15-155`](../../../../crates/litchi-pptx/src/opened/tests.rs#L15),
which calls `notes::load_snapshot` separately from the candidate. In
particular, retain coverage for:

- notes validation after slide identity and name checks
  (`opened/tests.rs:356-388`);
- Transitional and Strict MCE, mixed conformance, and exact proof versus raw
  fallback (`opened/tests.rs:390-470`);
- name/root/notes error precedence and foreign inventory fallback
  (`opened/tests.rs:472-521`);
- separate 16 MiB notes-root and 64 MiB part limits
  (`opened/tests.rs:524-565`); and
- independent package-source ownership and no cross-contamination
  (`opened/tests.rs:567-589`).

Add focused guards that instrument or otherwise count `materialize`/`own_blob`
for opened capture and assert zero notes payload-owning copies there, while an
explicit public notes snapshot still materializes its graph. Compare candidate
and legacy results for both successful notes graphs and every refusal, including
malformed XML, relationship metadata, duplicate names, aggregate limits and
orphan parts. Assert identical complete package revisions, part-digest
ownership, published bytes/readback and patch behavior. A reservation-failure
guard must still return the typed allocation error and leave no partial
snapshot if the validation-only metadata vector is made fallible.

Only after these differential guards pass should the 0829 workflow be rerun on
the current HEAD with matched before/after cold baselines. No native Office or
broad-producer benefit is implied by this source review.

## Expected implementation scope

The likely scope is `crates/litchi-pptx/src/opened/model.rs`,
`crates/litchi-pptx/src/notes/package.rs`, and a small private factoring of
validation/fingerprint helpers in `notes/transaction.rs` or
`notes/validation.rs`, plus focused opened-capture tests. It should not touch
Scene/raw-shape scanning, the writer, public notes graph semantics, package
fingerprint policy, signatures or publication. The candidate remains a
next-optimization consideration until current-head attribution demonstrates
that the avoided temporary work clears the project's acceptance floor.
