# 0361la — IWA package replay uses the shared commit triad

- Date: 2026-10-22.
- Base: `50122ae80aa68d9d80cde4d722db7b37fdcebb4f`.
- Contract: 0351 §1–§2.2; follow-up to 0361l.
- Status: source migration complete on `uo2/iwa-detection-api`; targeted
  validation and coordinator integration remain pending.

## Package owner

`litchi_iwa::Package` is the immutable host owner. Its raw entry and archive
mutators remain crate-private because the format-owned Pages, Numbers, and
Keynote editors are the public staging surface. A package edit now publishes
`litchi_core::edit::Commit<Package, Patch>` with the shared
`litchi_core::edit::Diagnostics`; `litchi_iwa::Patch` is the direct canonical
re-export of the source-checked ordered-entry patch owned by
`litchi-iwa-package`.

`Package::apply(&Patch)` validates the package, checks the patch against the
complete ordered entry content, and returns the new package, patch, and
content-free diagnostics. A patch therefore replays onto any package with the
same recorded base content, including independently parsed packages. No-op
replay retains exact ZIP source provenance and reports unchanged diagnostics;
changed replay drops that physical-source provenance because the raw package
layer has no reassembly proof for an arbitrary patch.

The previous public `Snapshot::apply`, `Snapshot::edit_with`, and local
snapshot-publishing `Commit<T>` are removed. They exposed a second owner and a
commit accessor for a snapshot rather than the package required by the
contract. Raw package capability is preserved by the public
`Package::edit() -> Edit` family: `Edit` exposes the bounded entry and archive
operations, and `Edit::commit()` publishes the same shared triad. Dropping an
edit discards its copy-on-write candidate; output is available from the
committed `Package`, while `Package::apply(&Patch)` is the replay family.
There is no generic callback publication wrapper or compatibility patch
spelling. The edit inherits the source package's checked `PackageLimits`; a
failed staged operation leaves both the edit candidate and source unchanged.

The stale-base path is crate-owned `Error::PatchConflict` with its fixed
content-free reason. `SourceChanged` remains a separate source freshness error.
The host `litchi-ops` facade has no dependency on `litchi-iwa`, so it does not
acquire an artificial dependency for an unreachable host error mapping; a
future adapter at that boundary must classify these errors as required by
0351 when it actually owns the dependency.

## Consumers and fidelity

Package tests now exercise clone-based copy-on-write reads, package-level
triad replay on an independently reopened matching-content package, inverse
replay, no-op byte preservation, stale conflict classification,
dropped/final-validation edit discard, atomic limit rejection, indexed reads,
and streaming output from the committed package. Existing format editors
continue to use the crate-private raw mutation seam and retain their prior
entry ordering, archive validation, source retention, and bounded output
behavior.

The 0361l detector migration and its consumers remain in the same branch and
record. No compatibility alias or deprecation is introduced.

## Validation status

Before this source migration, the detector library, IWA library, affected
consumer compilation, glossary unit tests, and the affected IWA tests passed.
The Keynote and Pages fuzz workspaces need their post-migration rerun. The
warnings-denied Clippy baseline remains open: current failures are the
pre-existing `soapberry-zip` dependency (14 `type_complexity`/
`too_many_arguments`), generated `litchi-iwa-protos` (32 generated/large-item
lints), `litchi-iwa-detect` (`type_complexity` in
`into_semantic_source`, `src/lib.rs:1093`), `litchi-iwa` (29 existing
argument/complexity lints), Pages (18), Keynote (35), and Numbers (40).
These are recorded as blockers rather than claimed as closed; Ops and facade
strict-Clippy runs were still pending when the host build hold began.
