# XLSX SVG batch optimization plan

Status: design only. This report records a bounded candidate and its validation
plan; it makes no production change and makes no performance claim about an
unmeasured candidate.

The evidence authority is the retained sealed baseline at
`docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/results/clean-ab954d91d/`.
It was measured from clean source `ab954d91dd6d6ca3e0df18b59ee826cef99bcae8`
with the lifecycle implementation pinned to
`ac288a303264f9ea0bb4081baa44031bee5b79a7`. The baseline passed the root and
independent verification lanes. The external `/var/tmp` copy is not used as
the design authority.

The design follows the optimization order and bounded-work requirements in
[`docs/GOAL.md`](../../GOAL.md), the immutable snapshot/edit/patch and exact
no-op rules in
[`ADR 0003`](../../adr/0003-snapshots-edits-and-patches.md), the finite
budget, bounded-cache, and measurement rules in
[`ADR 0005`](../../adr/0005-io-memory-and-performance.md), and the
signature, opaque-content, and lossless-preservation rules in
[`ADR 0006`](../../adr/0006-validation-security-and-compatibility.md).

## Hypothesis pending attribution

The leading bounded hypothesis is that a transaction-scoped SVG drawing
workset could turn the source drawing scan, source relationship view, final
owner projection, exact part/relationship budget, and one source-preserving XML
splice into shared facts. The workset would be consumed by both the candidate
budget check and the composed planner. It would keep a delta overlay for
relationships and materialize the full relationship/XML tokens only at the
existing source-preserving boundaries. This is a hypothesis to test after
public-phase attribution, not an implementation decision.

If attribution identifies SVG planning/proof as the dominant public phase, the
workset hypothesis predicts these duplicate-work reductions before any parser
or allocator change:

* The staging cache in `preflight_with_pending` currently retains only scalar
  owner and raster relationship facts. Commit planning rescans the same
  drawing in `project_svg_drawing_budget`, `new_drawing_plan_state`, and the
  batch planner. The candidate retains bounded source ranges and relationship
  facts as an owned index so the commit planner can prove the same operations
  without reparsing each phase.
* `preflight_candidate_budget` computes a drawing budget, then separately
  allocates media URIs, admits parts, plans the package-wide relationship
  topology, and plans content types. The candidate derives these metrics from
  the final projection once, while retaining the same final-state limits.
* The planning function `finalize_svg_target_cleanup` invokes its package-wide
  incoming-edge census even when no cleanup target exists. An attach-only
  workset can skip that *cleanup-planner* census when the target set is empty.
  Detach and replacement paths still perform the complete incoming-edge proof;
  the separate post-transition `validate_svg_final_removals` resource/closure
  validation remains required for every published SVG transition.
* `changed_relationship_ids` builds and sorts a combined before/after ID list.
  The workset can record the changed IDs as relationship deltas during the
  final projection, preserving deterministic ordering when the source-bound
  relationship edit is materialized.
* Final guards still require one source inventory and one target inventory per
  worksheet part URI and drawing ordinal guard identity. Their common context can be shared between the
  selected pictures rather than copied into every `SvgFinalState`; the exact
  source and target bytes/tokens remain retained for patch validation.

`OwnedXmlPart::update_elements` already batches selected picture edits into one
source scan and one output allocation. It is therefore not the first target.
The first measurement should test whether those earlier repeated scans, graph
census, relationship-map construction, and guard-context copies are material;
it must not assume that they dominate the public operation.

## Evidence and limits of the evidence

The retained uncertainty report gives these absolute allocator-instrumented
observations for the same-drawing attach lanes:

| Lane | Input bytes | Process-median elapsed observations (ns) | Requested allocation (bytes) | Peak live delta (bytes) |
| --- | ---: | --- | ---: | ---: |
| `multi_picture_same_drawing_16` | 4,126 | 13,131,873; 13,121,158; 13,423,924.5 | 30,926,141 | 653,817 |
| `multi_picture_same_drawing_64` | 4,876 | 139,445,375; 138,903,262.5; 139,107,033.5 | 213,800,728 | 1,247,097 |
| `multi_picture_same_drawing_256` | 7,692 | 1,984,573,621; 1,988,579,013; 1,995,196,422.5 | 2,463,741,687 | 4,608,267 |

The retained final report also records p50 elapsed values of 13,140,157 ns,
139,116,654 ns, and 1,988,611,713 ns for those lanes. These are absolute
observations from three fresh processes with two warmups and twenty measured
samples per process. They do not prove a speedup, an asymptotic scaling law, or
that any particular pass dominates. They justify investigating repeated work
and establish the fixtures and timed scope for the matched candidate run.

The baseline covers 19 exact production source paths, 69 acceptance lanes,
three processes, and 4,871 inputs. Its receipts are immutable evidence. A
candidate must be built at a separate source pin and measured with the same
19-path source identity, fixture bytes, limits, process schedule, timed
operation, semantic checks, and receipt schema. No baseline limit or test is to
be altered to improve a result.

## Current source path and repeated work

The following is a source-pinned code reading of `ac288a303`, not an
instrumented attribution:

| Phase | Current function(s) | Work repeated for a same-drawing batch |
| --- | --- | --- |
| Stage | `semantic::transaction::Edit::stage_svg_attach`, `stage_svg_detach`, `svg_lifecycle::preflight_with_pending`, `build_preflight_cache` | One bounded `SourceDrawing::scan_with_limits` per drawing, followed by per-intent raster fallback lookup and projected-owner checks. The retained `DrawingPreflight` contains scalar facts only. |
| Candidate budget | `preflight_candidate_budget`, `project_svg_drawing_budget` | A fresh drawing scan and relationship-reference count; final owner/attachment growth, payload lengths, removed edges, and relationship IDs are reconstructed. |
| Package admission | `planned_svg_part_uris`, `topology::admit_parts`, `plan_final_relationship_topology` | Part-name/size and package relationship owners are walked again; source relationship maps are cloned or canonicalized for the final topology. |
| Composition | `plan_composed`, `new_drawing_plan_state`, `plan_composed_attach_batch`, `plan_composed_detach_batch`, `plan_composed_mixed_batch`, `plan_composed_replacement` | The drawing XML is scanned again for the plan state and again by the selected batch branch. Attach validation and relationship-ID allocation are repeated against the same source picture facts. |
| Cleanup | `finalize_svg_target_cleanup`, `incoming_relationship_targets_with_states` | A package-wide incoming-target index is built even for an attach-only empty cleanup set. Detach/replacement needs this proof; attach-only does not. |
| Relationship materialization | `changed_relationship_ids`, `materialize_relationship_transition` | Before/after IDs are collected, sorted, deduplicated, and then source relationship edits are planned. A full relationship clone is retained even when a small delta would suffice during projection. |
| Final proof | `capture_final_guards`, `capture_final_inventory`, `capture_final_state_from_inventory`, `validate_guard_states` | Source and target inventories are shared per worksheet part URI and drawing ordinal guard identity, but each selected picture gets a separate full `SvgFinalState`, including repeated common relationship/content-type context clones. Patch replay redoes the inventory read, as required for stale detection. |
| Harness | `execute_multi_picture_attach` | The timed scope includes workbook open, every stage call, commit, plain serialization, output reopen, second serialization, and all semantic owner/graph/opaque checks. These harness scans are part of the sealed scope and are not removed from the candidate. |

The current planner is already better than one XML rewrite per picture: the
common attach branch performs one `OwnedXmlPart::update_elements` call for its
group. The remaining concern is the work around that splice. For a batch of
`N` selected pictures in one drawing, let `D` be drawing XML bytes/events, `R`
be its relationship members, `M` be package relationship owners/parts, `G` be
relationship graph identities, and `V` be final output bytes. The current
commit path performs several `O(D + R)` passes, several `O(M + G)` package
passes, and per-picture validation/guard construction in addition to the
irreducible `O(V)` output and `N` payload copies. The exact constants and
branch mix depend on detach, replacement, ordinary graph edits, and cleanup;
the baseline does not establish an asymptotic bound.

The falsifiable first target is narrower than an elapsed-time promise. For an
attach-only batch affecting one drawing, the commit planner should have one
shared source-drawing facts build for budget plus composition, rather than the
current budget, plan-state, and attach-batch scans. The staging preflight scan,
one source inventory and one target inventory per worksheet/drawing guard identity, and the
sealed harness's output validation remain separately counted. The attach-only
planner should also perform zero cleanup-planner incoming-target census work
when the cleanup set is empty. That target does not remove the one
post-transition `validate_svg_final_removals` pass, which still walks final
relationships to validate removed-target reachability and all resource
ceilings. The final relationship transition may materialize one
source-preserving token, and the harness must retain its existing
serialization/reopen count. These are static work-count targets to verify with
diagnostic counters; they are not claims that elapsed time or allocation will
improve by a corresponding factor.

Before selecting this target, collect public-phase attribution with the
separate exploratory profiler scaffold. The phases should at least distinguish
workbook open, public staging, SVG commit/planning, package publication and
serialization, output reopen/readback, and semantic validation. Where the
scaffold can do so without changing semantics, count drawing scans,
relationship-owner walks, XML/relationship materializations, and guard-state
clones inside the phase that performs them. The profiler is an attribution
instrument, not a replacement benchmark and not permission to change the
sealed baseline harness. Run the attribution against the clean source and
lifecycle production pin `ac288a303264f9ea0bb4081baa44031bee5b79a7`, with the
same `multi_picture_same_drawing_{16,64,256}` fixture bytes and public timed
scope used by the retained baseline. Report per-process phase medians and
ranges; with three processes, treat them as descriptive uncertainty rather
than a causal or speedup estimate.

Predeclare this plan's materiality gate before collecting attribution: classify
SVG planning/proof as material only when its inclusive median share is at least
10% of the public timed operation in at least one of the three same-drawing
lanes and the repeated scan/census counter shows at least two passes over the
same drawing or package facts. If the observed process-median range straddles
that threshold, classify attribution as indeterminate and defer this workset.
If another phase dominates, record this workset as a deferred hypothesis and
design the next candidate around the measured phase.

## Conditional workset design

The following names describe the design, not an API already present in the
tree. They become an implementation candidate only if public-phase attribution
supports the planner/proof hypothesis. The first implementation should keep
the type private to the workbook edit planner unless reuse by staging proves
necessary.

```text
SvgDrawingWorkset {
    physical_key: canonical physical drawing URI,
    physical_source_identity: exact drawing XML and drawing-rels source tokens,
    source_xml: Arc<Vec<u8>>,
    drawing_facts: bounded picture/range/owner/opaque index,
    before_relationships: source-bound relationship token and read-only view,
    relationship_delta: deterministic add/remove/replace overlay,
    projected_owners: final owner state in intent order,
    media_allocations: final URI, relationship ID, payload Arc for each attach,
    final_budget: exact final part/relationship/content-type metrics,
    worksheet_guard_identities: bounded worksheet-rel/picture guard contexts,
}
```

The physical drawing facts are keyed by canonical drawing-part URI and may be
shared by selectors that resolve to that same immutable physical part. The
worksheet guard identity is separate: it includes the worksheet URI, the
worksheet-to-drawing relationship source token and member-presence state, the
drawing ordinal, and the picture ordinal. A shared physical drawing must not
collapse these worksheet associations into one guard; a changed worksheet
relationship or member-presence token invalidates the affected guard even when
the drawing-part bytes are unchanged.

`drawing_facts` should retain byte ranges and compact offsets into
`source_xml`, not parsed nodes with unconstrained ownership. It must include
the picture selector, anchor, raster relationship ID, owner state, owner
extension range, dialect/prefix information needed by the existing splice,
embedded relationship reference counts, and the opaque/unsupported owner
classification. It also retains the exact drawing relationship members and
the target/type/mode facts needed by fallback validation. The source XML and
drawing relationship token are shared `Arc`/source-backed values, so retaining
the workset does not copy the source bytes for each picture. Worksheet
relationship tokens remain in the guard-context records rather than being
treated as physical drawing facts.

`relationship_delta` is an overlay keyed by relationship ID. It records the
final value (`Some`) or removal (`None`) and keeps the source map immutable
until `materialize_relationship_transition`. It must not hide unknown
relationships: all unchanged source members remain visible through the
source-backed view. If an existing API cannot prove that overlay semantics
preserve ordering/comments/lexical details, materialize one full
`OwnedRelationships` value at that boundary and keep the optimization at the
scan/clone level.

The materialization boundary is mandatory and source-preserving. Relationship
changes must continue through the existing source-bound relationship plan
(including its member-presence state), so an absent `.rels` member, an existing
member with comments/order, and an empty final relationship set retain the
same presence/removal behavior as the current implementation. Content-type
changes must likewise continue through the source-bound `OwnedContentTypes`
plan, preserving defaults, overrides, ordering, comments, and the final
`[Content_Types].xml` member. A map-to-canonical-XML serializer or a plan that
silently loses relationship/content-type member presence is outside this
hypothesis.

The workset is built once for each affected drawing. It then performs these
steps in order:

1. Resolve all intent selectors and coalesce repeated selectors using the
   existing `coalesce_repeated_group` rules. Preserve call order, replacement
   detection, and error order.
2. Validate every raster fallback and owner transition against the one
   `drawing_facts` index. Apply each intent to `projected_owners` in order;
   do not treat a cached owner as authority if the source identity check fails.
3. Allocate final media URIs and relationship IDs once, in the same order as
   the current `RelationshipIdAllocator` and part-prefix allocator. Store the
   `Arc<Vec<u8>>` payload already owned by the intent; do not create a second
   payload copy for budgeting or final media construction.
4. Produce the final drawing XML edits and relationship overlay. The existing
   `OwnedXmlPart::update_elements` and source relationship plan remain the
   source-preserving materialization boundaries.
5. Derive the exact final budget and topology from the projected state, then
   admit the operation. The budget result is reused by the planner; it is not
   recomputed by a second scanner.
6. Run incoming-edge cleanup only when the final cleanup target set is
   nonempty. For detach/replacement, retain the complete package-wide census
   and all ordinary graph/relationship overrides.
7. Hand the shared context for each worksheet/drawing guard identity to final guard capture. The
   target workbook still gets one fresh target inventory per guard identity after
   publication; replay validation still re-reads the package.

This is deliberately a removal of duplicate proof work, not an invitation to
skip a proof. The candidate may combine equivalent reads only when the same
immutable source token and exact final projection are still checked.

## Ownership, lifetime, and invalidation

The preferred ownership is transaction-scoped. A workset is reachable only
from the active `Edit`/SVG plan and is dropped on commit, abort, or an error.
It is never global, process-wide, snapshot-wide, or an unbounded LRU. The
number of worksets is at most the number of affected drawings, bounded by the
existing intent cap `MAX_SVG_LIFECYCLE_INTENTS` (65,536). Picture, node,
relationship, fragment, and retained-byte counts remain bounded by the full
caller `ReadLimits`, the existing `scan_limits(workbook)` projection, and the
existing intent/payload caps. Every new index must have an explicit bounded
cardinality or byte bound, use the existing checked reservation/error paths,
and be rejected when the applicable existing limit is exceeded. Do not invent
a second operation-budget accounting model or silently charge this index to a
different limit. A cache entry is not allowed to make a previously refused
input acceptable.

The physical source identity contains the immutable base lineage, canonical
drawing URI, exact source drawing bytes (or their source token plus a
length/digest used only for lookup), and the exact drawing relationship token.
Each worksheet guard identity additionally contains the worksheet XML and
worksheet-to-drawing relationship token/member-presence state. A digest may
avoid a lookup; it is not authorization. Before a workset is consumed, the
candidate must compare the exact source bytes/tokens required by the existing
source-bound patch contract. A hash-only match cannot authorize a patch.

Invalidate and rebuild the affected workset when any of these occur:

* a staged ordinary part, worksheet relationship, drawing relationship,
  content-type, or graph edit can change the drawing, its worksheet edge, a
  media target, or an incoming-edge decision;
* a source package `Arc`/byte token, relationship source token, read-limit
  set, or base snapshot lineage differs;
* `Edit::join` would combine transactions whose bases are not the same
  immutable snapshot: the existing gate is `Arc::ptr_eq` on the base workbook
  inner value, so byte-equal but separately allocated snapshots must still be
  rejected. For a same-`Arc` join, merge or retain a workset only when its
  physical key, worksheet guard identity, source tokens, and limits agree;
  otherwise invalidate and rebuild it after the existing disjointness checks;
* an unsupported/opaque relationship or owner is encountered outside the
  indexed ranges, or a parser/limit version changes the facts required by the
  splice.

Staging-cache updates must be transactional. Build a new facts entry and its
projected-owner delta in temporary state, then publish the cache entry,
projected map, intent, and staged-payload counter only after every validation,
allocation, and reservation for that staging call succeeds. An error must
leave the `Edit`'s prior cache/maps, intent vector, and byte counters unchanged.
The internal absent-owner detach result may retain immutable facts for a later
call, but it must not publish a projected-owner or lifecycle-intent mutation.
`Edit::join` must follow the same rule: its existing same-`Arc` and disjointness
checks, reservations, and combined staged-payload check complete before either
edit's cache state is merged; a rejected join leaves `self` unchanged.

The workset is an optimization hint. `SvgReadGuard`, source relationship
tokens, content-type tokens, final states, and ordinary graph guards remain
the authority for patch application, inverse, replay, and stale detection.
Any guard mismatch returns the existing `PatchConflict` path and discards the
workset; it must never fall back to a stale cached projection.

## Exact limits and budget semantics

The candidate must preserve the current limits and check the final projected
state. It must retain at least the following boundaries:

* `MAX_SVG_LIFECYCLE_INTENTS` (65,536), the per-input
  `MAX_SVG_INPUT_BYTES` bound (combined with `payload_limit`), the separate
  cumulative staged-payload pool checked by `check_payload_budget`,
  `MAX_DRAWING_BYTES`, `MAX_OUTPUT_BYTES`, `MAX_ID_ATTEMPTS`, and generated
  relationship-ID exhaustion behavior;
* caller `max_part_bytes`, `max_total_part_bytes`, `max_parts`,
  `max_relationships_per_part`, `max_total_relationships`,
  `max_relationship_parts`, `max_total_relationship_xml_bytes`,
  `max_total_relationship_xml_events`, `max_relationship_graph_nodes`,
  `max_content_type_mappings`, `max_content_types_bytes`, XML bytes/events,
  XML depth, picture count, relationship-reference count, and fragment bytes;
* the full caller `ReadLimits` at every new scanner/index/materialization
  boundary, plus the existing source/package read limits and staged input
  ownership check in `check_payload_budget`.

The per-input and staged-pool limits are distinct. `stage_svg_attach` rejects
one borrowed payload larger than `payload_limit(workbook).min(MAX_SVG_INPUT_BYTES)`
before copying it. `check_payload_budget` separately checks the cumulative
`svg_staged_payload_bytes + incoming` pool against caller
`max_total_part_bytes`; that temporary pool is not the final physical-part
budget and is not cancelled merely because a later composed edit detaches the
picture.

For a final projection, the candidate budget is defined as follows:

```text
N_attach = number of final owner transitions that materialize an embedded SVG
            (including a replacement whose final payload is the last attach)
payload_sum = sum of the final payload length for those transitions
P0 = map of canonical physical OPC part names from package.iter_parts()
     (relationship parts and [Content_Types].xml are excluded)
Pfinal = (P0 - final physical removals)
         union final ordinary physical additions
         union unique planned SVG media part names
final_parts = cardinality(Pfinal)
final_part_bytes = sum of final blob lengths over Pfinal
final_relationships = sum of member counts in every final relationship owner
                      (including unchanged source members and all final deltas)
final_relationship_parts = count of final .rels members with member presence
final_content_types = source manifest after final removals and additions
```

Replacing an existing physical drawing or worksheet part changes its final
blob length but does not add a second key to `Pfinal`; relationship-part and
content-type XML changes are accounted only in their respective final metrics.

The implementation must use checked arithmetic and reject overflow. The
candidate must check `final_parts` against `max_parts`, each final blob length
against `max_part_bytes`, and aggregate `final_part_bytes` against
`max_total_part_bytes`; relationship
member/part/XML/event/graph counts against their existing relationship limits;
and final content-type bytes/mappings against their existing manifest limits.
Relationship parts and `[Content_Types].xml` are never added to the physical
`iter_parts()` map. It must not add old and new values simultaneously when the
final projection replaces or cancels an entry. Conversely, it must not
under-count a shared edge: the cleanup-planner incoming-edge census and the
post-transition `validate_svg_final_removals` closure check remain required
where applicable before removing any media target, including targets shared
by pictures or ordinary graph owners. Staged payload bytes remain an
independent temporary transaction pool; the fact that a later detach cancels
an attach must not bypass ingress bounds.

## Semantic invariants that cannot change

The optimization is accepted only if the following behavior is byte- and
error-equivalent to the source-pinned implementation for the covered inputs.

**No-op and refusal.** At the internal
`Edit::stage_svg_detach` seam, a picture with no admitted embedded owner
continues to return `Ok(false)` without appending an intent or publishing a
part, relationship, media, or manifest change. The public
`WorksheetEdit::detach_svg` API deliberately discards that boolean and returns
`Result<&mut Self>` for chaining, so a public no-op is observed as a successful
chain result that appends no SVG intent and leaves the SVG/package state
unchanged at commit; the public API must not be documented as returning
`Ok(false)`. A transaction with no effective lifecycle change returns the
original snapshot/package allocations where the existing edit contract promises
exact sharing. Existing embedded-owner
refusal, linked/opaque/ambiguous/refused owner errors, invalid selector errors,
missing raster errors, unsupported relationship errors, and limit failures
retain their current type and precondition order. Coalescing may avoid work
only after those same staged operations would have been admitted and their
observable error order is preserved.

**Inverse and stale patches.** A committed `Patch` remains reversible and
source-bound. Its before/after XML, relationship, content-type, media, opaque
bytes, and `SvgReadGuard`/`SvgFinalGuard` expectations remain exact. Applying
the inverse or replaying a patch on a changed drawing, relationship member,
worksheet edge, content type, or shared SVG media part must re-read and reject
the preimage with `PatchConflict`. A workset must not be serialized as patch
authority and must not be reused across a changed base snapshot.

**Signatures.** `Edit::new` continues to enforce the existing unsigned/edit
boundary through `codec::ensure_unsigned`. The candidate cannot cache around
signature verification, mutate signed bytes, or turn a signed-source refusal
into a planner success. Any signature-related source identity or guard remains
part of the existing error path.

**Opaque and lossless content.** The workset indexes opaque extensions and
unknown descendants only to locate and protect them. It never canonicalizes
the whole drawing, drops unknown relationships, rewrites namespace bindings,
or normalizes untouched XML. Untouched worksheet/drawing XML, relationship
members, content-type ordering, comments, lexical details, compression,
timestamps, and unrelated package parts continue through the existing
source-preserving token/materialization APIs. Unsupported or ambiguous owner
states remain typed refusals or preserved opaque states exactly as before.

**Composed edits.** Distinct selectors in one drawing use one final projection
and one XML splice where the current branch does so. Repeated selectors use
the current coalescing and replacement rules. Mixed attach/detach sequences
retain call order, deterministic media/relationship allocation, collision
behavior, cleanup decisions, and ordinary graph conflict checks. Attach,
detach, replacement, and mixed batches each retain their required source
scanner limits and fallback validation. A shared raster relationship or SVG
media target is removed only after the final incoming-edge census proves that
no surviving edge references it.

**Readback and guards.** The timed harness continues to reopen the serialized
output and run all existing owner, relationship, graph, media, content-type,
and opaque checks. Final guard capture still proves the raster anchor and
fallback are unchanged, the expected embedded SVG state is present/absent,
the SVG target is inert `image/svg+xml`, and the manifest covers it. Sharing a
common guard context may reduce copies; it cannot remove per-selector expected
state or the final package read.

## Implementation boundary and coordination

The likely implementation touches `svg_lifecycle.rs` and, if stage-time facts
are reused, the SVG fields and transaction flow in
`workbook/edit/semantic/transaction.rs`. If final guard context ownership is
changed, it may touch `workbook/edit/model.rs` and the commit/readback seam.
Those are shared edit/model/semantic transaction files. The active pivot coder
must review the ownership and invalidation design before any production patch
is started. This report intentionally makes no such patch.

If attribution confirms planner/proof work as the dominant phase, a low-risk
first slice would keep the workset local to `svg_lifecycle::plan` and merge
`preflight_candidate_budget` with `plan_composed` around one per-drawing
source index. That slice could remove the commit-time budget/state/batch
rescans and the empty attach cleanup census without changing the public `Edit`
model. A second slice could replace repeated final-guard context clones with
`Arc` sharing after the first slice passes all patch and replay tests. A
stage-to-commit cache should be attempted only if source identity and
invalidation can be made explicit; otherwise its cross-phase lifetime adds
risk without being necessary for the first measurement. These are conditional
implementation hypotheses, not a request to make the broad workset change
before attribution.

Do not begin with a new parallel executor, a global cache, a whole-drawing
canonical serializer, or a hand-optimized XML parser. None is needed to test
whether repeated proof work explains a material portion of the measured
workload, and each would enlarge the semantic surface before evidence exists.

## Required tests before measurement

The candidate implementation, when authorized, must run the existing unit,
integration, root-verifier, and independent-review checks without modifying
fixtures or limits. In addition, add focused regression coverage for the
workset boundaries (the tests are a later implementation task, not part of
this design-only change):

1. Attach 16, 64, and 256 pictures in one drawing and assert every anchor,
   raster edge, SVG owner, media byte, relationship type/mode/target, content
   type, opaque fragment, and output reopen result.
2. Attach to an existing owner, detach an absent owner, repeated attach/detach,
   detach-then-attach replacement, mixed selectors, and duplicate selectors;
   assert the internal no-op boolean/public chaining distinction, exact no-op
   sharing, error order, and final projection.
3. Shared and distinct SVG media targets, incoming edges from ordinary graph
   edits, relationship-ID collisions, part-name collisions, and cleanup
   cancellation; assert no referenced part is removed and final counts are
   bounded by the limits.
4. Namespace-heavy, strict-namespace, malformed/duplicate/linked/opaque owner,
   unsupported fallback, missing relationship/media, and XML/relationship
   limit cases; assert the same refusal class and no partial publication.
5. Create a patch, apply it, apply its inverse, and replay it after changing
   drawing XML, relationship XML, content types, or shared media; assert exact
   bytes on success and `PatchConflict` on stale input. Include signed-source
   refusal through the existing edit boundary.
6. Join disjoint edits from the same base `Arc`; assert transactional cache
   merge. Join byte-equal but separately allocated snapshots and different
   base `Arc`s; assert the existing `DifferentSnapshot` refusal and no cache
   reuse.

For each successful and refused case, compare the candidate against the
baseline on serialized package bytes where the existing contract requires
identity, semantic inventory, patch guards, allocation validity, and all
reported limits. A performance result cannot waive a semantic mismatch.

## Matched measurement plan

The sealed baseline remains immutable historical evidence and is never edited,
regenerated in place, or re-pinned. It can be a direct reference arm only
after comparing source manifests and transitive capture to the candidate. If
the intervening pivot changes shared or transitive production sources, the
sealed `ab954d91d` result must not be treated as the causal control for the
candidate.

For a causal comparison after any such intervening change, create a fresh
matched pair: a clean baseline checkout at the common pre-optimization parent
and a clean candidate checkout at that same parent plus only the approved
optimization. Record separate source/production pins, exact 19-path
production manifests, and identical transitive dependency capture for both
arms. Keep the sealed bundle as the immutable historical reference. If there
are no intervening relevant changes, a candidate checkout may instead be
compared with the sealed arm, but it still needs its own recorded pin and
source-manifest verification.

Every causal pair must use the same committed fixture corpus, harness and
transitive capture, crate/profile/toolchain, host class where available,
caller limits, `multi_picture_same_drawing_{16,64,256}` fixture construction,
and 69 acceptance lanes. The public-phase profiler is a separate exploratory
scaffold and must not silently alter one arm's timed scope. No result may
combine receipts from different timed scopes or from mismatched transitive
source captures.

The sealed `execute_multi_picture_attach` timed interval remains unchanged:

1. Build the fixture outside the timer.
2. Start a fresh `Workbook` from fixture bytes inside the timer.
3. Stage one attach per selected picture through the public API.
4. Commit, publish, serialize with `to_plain_bytes`, reopen the output, and
   serialize/read it again.
5. Run the existing all-owner, relationship, media, graph, manifest, and
   opaque semantic checks before stopping the timer and taking the allocator
   receipt.

Use three fresh processes, two warmups, and twenty measured samples per
process, preserving the current receipt fields and allocator accounting on
both arms of a causal pair.
Baseline `stage_ns`/source I/O fields are null; do not silently introduce a
candidate-only stage boundary and compare it as if it were a baseline stage.
If implementation counters are useful, collect them in a separate diagnostic
run or add the same optional fields to both separately generated runs without
changing the timed operation. The primary comparison remains elapsed time,
requested/direct/reallocation bytes, peak live delta, semantic success, output
exactness, and allocation validity under the unchanged scope.

Report p50/p95/p99 and all process medians for 16/64/256, plus every acceptance
lane and verifier result. Preserve the uncertainty framing: report absolute
observations and confidence/dispersion information, then state whether the
candidate is worth further work. Do not convert the three fixture sizes into a
speedup or scaling claim without a larger controlled study. A candidate target
is material for this bounded investigation only if, predeclared before the
matched run, it lowers p50 elapsed time or requested allocation by at least 10%
in one same-drawing lane on all three process medians, has no semantic or
limit regression, and does not worsen another same-drawing lane by more than
the baseline process-median range. Overlapping process-median ranges are
reported as inconclusive. This is a go/no-go criterion for this investigation,
not a general speedup or scaling claim; the document sets the target but does
not assert that it has been met.

## Decision gate

The first gate is public-phase attribution from the separate exploratory
profiler scaffold. It must identify the dominant phase and establish whether
SVG planning/proof work, rather than open, publish/serialization,
reopen/readback, or semantic validation, is large enough to justify this
hypothesis. The scaffold must not change the sealed baseline or be used to
claim a candidate speedup.

Only if that attribution supports the planner/proof hypothesis should the
active pivot coder review the transaction ownership and invalidation boundary.
At that point the local workset and empty-cleanup fast path are candidate
experiments, not a preselected broad implementation. Keep the sealed 19-path
evidence immutable; choose the direct sealed comparison or a fresh matched
baseline/candidate pair according to the transitive-source rule above. Revert
the candidate if it changes no-op sharing, public chaining behavior,
source-bound refusal, inverse/stale behavior, opaque preservation, final
budgets, or composed-edit ordering. Low-level parser or allocator work is
justified only after attribution and matched evidence identify a remaining
dominant cost.
