# 0916 DOCX section-layout target reuse design review

## Finding

This review is design-only; it makes no production-code change and makes no
performance claim. The commit-only route currently calls
`publish_section_layout_patch_to_stream` (`crates/litchi-docx/src/source_backed.rs:1271`),
which captures and parses the current main-document snapshot at
`main_section_layout_snapshot` and then calls `Patch::apply`
(`crates/litchi-docx/src/section/layout.rs:547`). For a changed patch,
`Patch::apply` validates the retained `after` bytes again through
`Snapshot::from_shared_xml`. `Edit::commit` already retains that validated
target snapshot in `Commit` (`layout.rs:375`), so this second target inventory
and section scan is the highest-impact avoidable work in the reviewed route.
The retained profile attributes 6.38M guest instructions to this timed
`Patch::apply` target parse within 240.6M guest instructions for the whole
publication; these are prioritization data only, not a native speed claim.

The public patch route must keep reconstructing its target: a `Patch` does not
retain a parsed snapshot and can be applied to any otherwise authorized
`Snapshot`. The candidate is therefore limited to
`publish_section_layout_commit_to_stream`; it does not add a parsed cache to
`Patch`, alter `Patch::apply`, borrow another resolver, or weaken any source
fingerprint.

## Implementable shape

1. Extract the first three stages of `Patch::apply` into a private or
   `pub(crate)` `authorize_source(&self, source: &Snapshot) -> Result<()>`:
   exact `before` bytes, the existing `PatchAuthorization` match, and the
   existing `limits` equality check, in that order and with the same typed
   errors. Keep `Patch::apply` calling this helper, then keep its exact no-op
   `Ok(source.clone())` and its changed `Snapshot::from_shared_xml` path.
2. Add a private snapshot helper that clones the retained XML/inventory Arcs
   and copies the current source version, lineage, artifact fingerprint, and
   limits. This is identity rebinding only; it parses nothing. A private
   `Commit::apply_to_current(&self, current: &Snapshot) -> Result<Snapshot>`
   should first call `self.patch.authorize_source(current)`, return
   `current.clone()` for a no-op, and otherwise return the retained commit
   target through that identity helper. The commit and patch are private and
   are constructed together, so their target XML invariant remains internal.
3. Make `publish_section_layout_commit_to_stream` capture the current snapshot
   exactly as today, call `commit.apply_to_current(&current)`, and retain the
   existing overlay/no-op output, output fingerprint, inverse reauthorization,
   and `Publication::new` code. Preserve the delegated operation string
   `"publish_section_layout_patch_to_stream"` when calling
   `main_section_layout_snapshot`; changing it would change typed refusal
   context. Leave `publish_section_layout_patch_to_stream` on `Patch::apply`
   so the public patch contract remains unchanged.

The current capture still performs `check_execution`, main-part/content-type
validation, the owned-source restriction, `source_version`, main XML load,
complete `SourceArtifact::fingerprint`, target section parsing for the current
source, the post-capture source-version fence, and the final execution check.
The helper is reached only after those checks. The existing OPC writer still
owns signature refusal, cancellation/work-budget checks, monitored source
reads, preservation planning, sink accounting, and the final freshness fence.

## Required invariants and risks

- Never select `commit.snapshot()` before authorization. A foreign, stale, or
  mutated package must still fail before output.
- Preserve the authorization order exactly: source XML mismatch gives
  `SectionLayoutStaleSource`; then detached/foreign/revision/fingerprint
  authorization errors; then `SectionLayoutPolicyConflict`.
- A no-op must use `current.clone()`, matching `Patch::apply` exactly. Using
  the commit's retained snapshot would return the edit-time identity and could
  change lineage/version/limit or parsed-allocation semantics.
- A changed target must inherit the authorized current identity before the
  existing `with_artifact_fingerprint` publication step. Do not use the
  commit's edit-time identity as an authorization shortcut.
- Keep the complete artifact fingerprint. Equal main XML from a different ZIP
  artifact must still reach the existing lineage/version/fingerprint checks.
- Reusing the target is safe because `Edit::commit` has already performed
  replacement construction, limit checks, topology validation, and semantic
  readback, and `Commit` exposes no mutation or alternate constructor. Do not
  generalize this into a parsed cache or retain any additional target beyond
  the existing `Commit` lifetime.
- Preserve signed no-op and changed-source behavior, partial-sink behavior,
  output bytes, inverse authorization, and all source mutation fences. The
  optimization removes only the duplicate in-memory target parse.

## Verification matrix

The existing `source_backed_sections` corpus remains the primary oracle:
`source_backed_existing_section_layout_is_exact_and_cell_safe`, exact complex
no-op, strict/default namespace, foreign children/comments/barriers, table-cell
exclusion, `section_layout_signed_noop_changed_and_partial_sink_are_typed`,
and `section_layout_detached_patch_cannot_publish_to_source_package` remain
the unchanged publication and preservation coverage. The existing section
input/event/depth/bytes limits and output-preservation assertions also remain
in that corpus.

The candidate adds focused coverage for the optimization itself:

- the internal `Commit::apply_to_current` test checks changed XML/inventory
  allocation reuse and exact no-op current allocation reuse;
- the existing changed publication test now checks retained inventory
  allocation sharing;
- `section_layout_commit_rechecks_complete_artifact_before_target_reuse`
  checks a complete-fingerprint refusal before output; and
- `section_layout_commit_rejects_revision_change_before_output` checks the
  source freshness fence and empty output.

The extracted `authorize_source` body was reviewed as an unchanged move from
`Patch::apply`; its stale, foreign, detached, and public no-op behavior remains
covered by the existing tests, while the new publication tests exercise the
commit route's foreign, detached, and fingerprint guards. Exhaustive duplicate
tests for every unchanged typed-error precedence branch were not added. The
writer's signature, cancellation, resource-limit, partial-sink, and final
freshness behavior is unchanged and was verified by the diff against the
existing corpus rather than by new optimization-specific tests.

The frozen change-0916 native protocol measures only this commit-only target
reuse candidate with the existing 0914 corpus and gates adoption at the
predeclared eight-pair median ratio of at least 1.05 with a lower confidence
bound above 1. The rejected 0915 resolver-borrow experiment (1.014 median)
does not authorize this change or provide evidence for it.

## Candidate review and ADR fit

The read-only candidate under `/home/zhuhe/litchi-performance-0916/candidate/source/tree`
matches the design in the authorization and target-reuse paths. Its public
`Patch::apply` path still parses changed targets, while the private commit path
rebinds the retained target to the freshly authorized current identity. The
candidate also retains the complete artifact fingerprint check and the original
operation string.

The candidate initially dropped one of the two baseline
`trim_clean_paragraph_index_for_operation` calls; its current revision restores
the second call immediately before snapshot capture, preserving the shared
paragraph-cache behavior. Its focused layout unit test now checks both changed
target XML/inventory allocation reuse and the no-op `current.clone()` XML and
inventory allocation. The source-backed additions cover complete-artifact and
revision freshness before output.

The design is compatible with ADR 0003's immutable commit/patch model, ADR
0005's eliminate-work priority and exact-source authorization, ADR 0006's
preservation and typed-refusal rules, and ADRs 0010, 0011, and 0024's crate and
physical-package ownership. ADR 0032 is not implicated: no new memo or cache is
introduced; the already-retained commit target is reused only within the
commit publication operation. The only implementation prerequisite is a
private identity-rebinding helper plus internal tests for the invariants above;
there is no accepted-ADR blocker.
