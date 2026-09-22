# 0730 DOC validated-render handoff design review

Status: read-only source qualification, 2026-09-22. This review owns this
file only. It makes no production, Cargo, native, probe, or measurement
change. The implementation decision below is for a bounded pilot; it is not a
production speedup or memory claim.

## Authority and unchanged constraints

The reviewed path is the ordinary public DOC body transaction in the current
worktree:

* `docs/GOAL.md` and the lossless, failure-atomic, no-hidden-cache
  requirements;
* `docs/performance/results/change-0729/constraints.json`, whose every listed
  `GOAL`, CRUD, ADR 0001 through ADR 0031, and ADR README hash matches the
  current file. The 0729 constraints map is unchanged;
* `docs/performance/results/change-0729/result-review.md`, which identifies the
  duplicate common render and requires identity, atomicity, one-shot
  ownership, an explicit retained-output bound, independent final validation,
  and exact no-op/preservation behavior; and
* the accepted retained-state amendment in ADR 0005 (2026-09-16,
  lines 2327-2377).

ADR 0005 is controlling for this pilot. Serialized state held by a public
value after the producing operation returns is retention policy, not an
invisible cache. Its ceiling belongs to the operation's finite limit policy
and is intersected with every other policy member when policies meet. The
owner must expose the held weight and a release operation. The retained bytes
must be the allocation already made by the operation; a new copy is not
permitted. Exceeding the ceiling falls back silently to recomputation when
recomputation is available, and the value must observably report that it holds
nothing. A retained document cannot spill to scratch storage.

The 0729 packet is descriptive attribution evidence. It shows that public
DOC `Finish` is 7.68--8.18% of the large fixture and 13.00--13.30% of the
small fixture, but its candidate recommendation is the only basis for this
review: a private DOC-owner batched handoff at the existing common-editor
boundary. The current source and semantic oracles remain authoritative.

## Current public path

The complete changed route is currently:

```text
Snapshot::open_bounded
  -> RevisionEditor::open_with_ole_file
  -> independent public Package::validate_ole_file
  -> exact source Arc retention
Snapshot::edit / Edit::new
  -> second strict RevisionEditor open from source bytes
Edit mutation
  -> public dependency and transaction-limit checks
  -> clone-first RevisionEditor mutation
  -> RevisionEditor::commit
  -> ObjectEditor candidate package check
  -> Reuse copy-through or layout render
  -> CFB reopen, package recapture, equal-stream allocation reuse,
     target discovery
  -> rendered Vec is currently dropped by the owner
Edit::commit
  -> RevisionEditor::finish (a second common render today)
  -> Snapshot::open_bounded (independent strict owner and public reader)
  -> source-checked Patch construction
```

The common candidate reopen is already a complete common validation boundary:
it checks the candidate package, renders under the current sector policy,
reopens the CFB, recaptures the package, reconciles equal stream allocations,
and rediscovers targets before publishing the candidate. It is not a public
DOC validation and cannot replace the final `Snapshot::open_bounded`.

## Mutation and no-op inventory

### Public `body_text::Edit`

| Entry point | Candidate operation | No-op or failure behavior | Handoff consequence |
| --- | --- | --- | --- |
| `replace_paragraph`, `replace_text` | Main-story length change calls `replace_plain_text`; equal-length non-main Unicode calls `replace_unicode_text_same_length` | Resolves target, rejects drawings/structure/tracked text/dependency and limit violations before mutation; equal text returns before the editor | A successful operation replaces the current token with the candidate token. Equal text preserves the current token. |
| `set_paragraph_bold`, `set_character_property` | `set_character_property_override`, then the corresponding CHPX rewrite and common commit | Same property returns before mutation; dependency and operation checks precede the candidate | A successful format change gets a new token. |
| `apply_transfer` | Delegates to `replace_text` after exact source-lineage check | The `Lineage` check is byte equality, not allocation identity; a plan from equal bytes is accepted even when its source policy differs | Intersect the plan's planning-receiver policy with the actual `Edit` policy before limit checks and mutation. |
| `apply_embedded_transfer` | Delegates to `add_embedded_object` | Same byte-equal lineage behavior; foreign bytes conflict before nested work | Apply the same planning-receiver policy meet; nested reopen then uses the effective bound. |
| `apply_picture_transfer`, private `install_picture` | `replace_with_picture_graph`; changes WordDocument, Table, and Data | Graph, drawing, tracked, formatting, position, and operation checks precede clone-first mutation | Apply the plan/receiver policy meet; Data is marked changed and the final batched render is retainable only if the Data topology is handled atomically. |
| private `restore_picture_text` and durable/inverse picture replay | `replace_picture_graph_with_text`; truncates Data and rewrites WordDocument/Table | Exact installed graph and append-only Data-tail preconditions; failures leave the current editor | Same Data-topology requirement. |
| `add_text`, revision replay, `dispose_revision` | `RevisionEditor::add_text`, `add`, `update`, `remove`, `accept`, `reject`, or `delete_revision_text` | Invalid revision/index/destructive dependency checks precede clone-first mutation | Each successful revision mutation gets a new token; a failed candidate leaves the old token. |
| `add_embedded_object` | Finishes a clone, opens the embedded-object transaction, commits it, then assigns a newly opened `RevisionEditor` | Collision and operation checks happen around a nested transaction; assignment occurs only after nested commit | `editor.clone().finish()` is deliberately a recomputation under clone-clears-token semantics; the assignment must reapply the retention limit and has no prior token. |
| `remove_embedded_object` | Same nested path, removing the managed field, preview, and ObjectPool storage | Missing resource and nested commit errors leave the outer editor unchanged | Reopen at the assignment site with the source transaction retention bound. |
| `set_embedded_display_as_icon` | Same nested path, changing OLEDS display metadata | Equal bit is an exact no-op; otherwise nested commit and reopen | Equal bit preserves a token; a changed nested assignment clears it and reapplies the bound. |
| `rollback` | Consumes the edit and returns the source snapshot | All staged state and any retained token are dropped | No retained serialized state is transferred to `Snapshot`. |
| `commit` / profiled commit | Consumes the `Edit`, obtains final bytes, then opens a changed result through strict owner plus independent public reader | Exact source bytes share the source snapshot; changed bytes must pass final validation before `Patch::new` | Consume a matching one-shot token in `RevisionEditor::finish`; do not skip final validation or patch construction. |

`Edit` has no `Clone` implementation. Keep it that way for the pilot. A
`Snapshot::clone` shares only the immutable source allocation and copies its
policies; it must never carry a render token.

### `tracked_revision::RevisionEditor`

Every package-changing method is clone-first: it builds a candidate, mutates
the candidate's WordDocument/Table/Data and parsed indexes, calls the private
`commit`, and assigns `*self = candidate` only after the common editor has
validated the candidate. The complete set is:

* `replace_plain_text`;
* `set_character_bold_override`, `set_character_italic_override`, and
  `set_character_underline_override`, through `set_character_override`;
* `replace_unicode_text_same_length`;
* `replace_with_picture_graph` and `replace_picture_graph_with_text`;
* `add_text`, `add`, `update`, and `remove`;
* `accept` and `reject`, which dispatch either to `remove`,
  `delete_revision_text`, or a candidate formatting rejection; and
* the private `delete_revision_text` path.

Parsing, `revisions`, author lookup, range checks, SPRM transformations,
CLX/CHPX/PAPX rewrites, FIB size updates, and dependency checks mutate only a
candidate's in-memory semantic state. They do not publish package state until
`commit` succeeds.

`RevisionEditor::commit` always replaces WordDocument and the selected Table.
It replaces Data when `data_changed` and Data already exists; when Data is
absent it currently calls `ObjectEditor::add_stream` and then performs the
replacement batch. That missing-Data branch is the only DOC commit topology
that currently performs an intermediate common render. Picture insertion and
picture reversal are the public routes that set `data_changed`. The final
rendered `Vec` from the replacement batch is still a complete rendering of
the clone's post-add state and is safe to retain; the intermediate render is
only an extra cost and does not make the final handoff stale.

`RevisionEditor::finish` is the final owner render. An exact unchanged editor
returns the source-backed result through the common editor. A changed editor
uses the common editor's default `SectorLayoutPolicy::Reuse`, including its
copy-through eligibility and layout-writer fallback. There is no
`RevisionEditor` policy setter today; direct tracked-revision callers use a
zero retention ceiling unless the DOC body owner opts in.

### Common `object::Editor`

The common editor's mutating paths all use an isolated clone and publish only
after `commit_candidate` (or its rendered variant) succeeds:

* `update` obtains a selected object and delegates to `replace`;
* `replace` captures a validated nested CFB, edits the candidate package, and
  commits it;
* `put_stream`, `put_stream_shared`, and the existing
  `put_stream_shared_with_rendered` replace one stream;
* `put_streams_shared` and the candidate
  `put_streams_shared_with_rendered` batch existing stream replacements;
* `add_stream` adds a new stream;
* `remove_stream` and `remove_streams` remove one or several streams;
* `update_link` edits an OLEDS stream through `put_stream`;
* `add_storage` and `remove_storage` edit selected compound objects; and
* `set_sector_layout_policy` changes only the future render policy and does
  not render or validate immediately.

`stream`, `stream_shared`, `snapshot`, `targets`, `objects`, and
`is_changed` do not mutate the package. `finish` is consuming: changed values
try Reuse copy-through and then the layout writer; unchanged values return the
exact original bytes. `commit` snapshots and finishes the editor.

The candidate `put_streams_shared_with_rendered` seam is the right common
boundary. It applies all existing-stream replacements to one clone, renders
and reopens once, publishes that candidate, and moves the validated `Vec<u8>`
to the caller. An empty or all-equal batch returns `Ok(None)` without cloning
or rendering. It is a handoff, not common-editor state and not a cache.

## Smallest correct handoff

The pilot should use a DOC-owned one-shot retained `Vec<u8>` and the existing
common rendered batch. The concrete shape is:

1. A one-render claim requires a narrow common batch variant that can apply
   one optional stream add together with the existing replacements before the
   single `commit_candidate_with_rendered` call. Its operation set is limited
   to the DOC commit's WordDocument, selected Table, and optional Data paths.
   It must preserve the existing stream-count, stream-size, package-check,
   Reuse, CFB-reopen, recapture, allocation-reuse, and discovery behavior. If
   adding this variant is deferred, run the current `add_stream` plus
   replacement route, retain the *final* replacement-batch `Vec` when its
   capacity is within the bound, and report that missing-Data candidates still
   perform two common renders. The intermediate render belongs only to the
   clone; it is not handed off or used as the final token.
2. Have `RevisionEditor::commit` receive the final rendered `Vec<u8>` from
   that batch and publish the candidate editor and token together. The
   existing current editor must remain unchanged if any package check, render,
   reopen, recapture, allocation reconciliation, or discovery step fails.
3. Keep a private `RetainedRender(Option<Vec<u8>>)` in `RevisionEditor` and a
   crate-private retention limit. The `RevisionEditor::Clone` implementation
   copies semantic/package state and the limit but always clones the token as
   `None`. This is a deliberate manual clone contract: a candidate cannot
   inherit a serialized result for the predecessor state, and cloning never
   copies or aliases caller-owned retained bytes. The old editor's token stays
   alive until candidate assignment or failure; that old-token-plus-new-render
   peak is unavoidable for clone-first mutation and must be measured.
4. On a successful common batch, retain the returned allocation only when its
   `Vec::capacity()` is at or below the configured ceiling. Move that exact
   `Vec` into the token; do not convert it to `Arc`, clone it, or serialize it
   again. If capacity exceeds the ceiling, drop the returned Vec and leave the
   token empty. This is silent recomputation fallback, not `Error::Refused`.
5. `RevisionEditor::finish` first takes a current token and returns that exact
   allocation. If no token is present it calls the existing common finish
   route. The token is consumed once; it is never installed on `ObjectEditor`,
   `Snapshot`, `Patch`, or a process-wide structure.

The token should carry a private generation and the render policy used to make
it, or equivalent owner-local freshness proof. Increment the generation on
every successful package publication, clear a token before any new package
candidate, and reject/clear a token if a future policy or limit setter can
change the render context. The current `RevisionEditor` encapsulates the
common editor and has no policy setter, so the clone-clears-token rule plus the
single publication boundary is the minimal identity proof today. Do not add a
public token constructor or a path/value-only key. The proof must cover the
source/editor instance, current package generation, target catalog, all DOC
replacement streams including Data topology, current sector policy, package
limits, and transaction retention limit.

The existing candidate render already validates the common package. The outer
`Edit::commit` must still consume the token only to obtain bytes, then run
`Snapshot::open_bounded` with strict owner validation, independent public
reader validation, exact source retention, and reversible patch construction.
The token is therefore a duplicate-render handoff, not a semantic-validation
shortcut.

## Transaction retention policy and policy meets

`TransactionLimits` currently has three private fields and a three-argument
`new` constructor. Preserve that constructor and add the retention member
additively:

```text
max_retained_render_bytes: usize
with_max_retained_render_bytes(self, maximum: usize) -> Self
max_retained_render_bytes(self) -> usize
```

Zero is a valid explicit off switch: the edit remains supported and simply
recomputes. The 0730 candidate may use an 8 MiB `Default` ceiling as a
provisional measurement setting. This value is evidence-gated and is not an
adopted production memory promise. `RevisionEditor::open` itself remains at a
zero ceiling so direct tracked-revision callers do not silently acquire
retention. `Snapshot::editor`/`Edit::new` opts the DOC body editor into the
snapshot's transaction ceiling. The three existing operation/replacement
limits and this fourth retention member are all observable through getters
and are included in `Debug`/equality by the existing derived implementations.

Add a componentwise `intersect` operation for the four transaction members:

```text
effective.operations                 = min(left.operations, right.operations)
effective.replacement_units         = min(left.replacement_units, right.replacement_units)
effective.total_units                = min(left.total_units, right.total_units)
effective.max_retained_render_bytes = min(left.max_retained_render_bytes,
                                           right.max_retained_render_bytes)
```

Every future policy meet must use this operation, never the more permissive
bound. The concrete current meet that needs handling is `body_text::Patch::apply`:
the exact source check may succeed while the supplied source snapshot and the
patch's before/after snapshots carry different `TransactionLimits`. Compute
the effective intersection of the supplied source, patch-before, and
patch-after policies before returning either the source clone for a no-op or
the after clone for a changed patch. Preserve the existing source/after byte
allocation; changing the policy metadata must not render or copy the
artifact. The same rule applies to inverse patch application.

`Lineage` is not allocation identity: `Lineage(Arc<[u8]>)` derives
`PartialEq`/`Eq`, so two separately opened snapshots with equal bytes compare
as the same lineage even when their `TransactionLimits` differ. The generic
`JoinedSubEdits` layer checks that byte-equal lineage and its own
`CompositionLimits`; it cannot see the DOC transaction policy. Consequently,
`Composition` is a real transaction-policy meet, not a one-policy shortcut.
Each `PreparedEdit` should carry the `TransactionLimits` of the snapshot that
validated it. `Composition` starts with its source policy, intersects the
incoming prepared policy after a successful semantic join, and commits from a
source snapshot carrying the effective componentwise minimum. A rejected join
must not change that effective policy. This covers independently reopened
equal-byte sources and makes the retained-render ceiling obey ADR 0005 along
with the three existing members.

Transfer plans (`TransferPlan`, `EmbeddedTransferPlan`, and
`PictureTransferPlan`) use the same byte-equal `Lineage`. Their donor policy
does not participate: the donor is read-only for this operation and contributes
semantic text or a bounded dependency closure, not a retained serialized
render. The *planning receiver* policy does participate and must be stored in
each plan. Before applying a plan, intersect that policy with the actual
receiving `Edit` policy for all four members, then run operation/replacement
checks under the effective policy. For example, a plan validated against an
equal-byte receiver with a zero retained-render ceiling must still produce no
retained render when applied to an equal-byte receiver opened with an 8 MiB
ceiling. If the lower operation bound is already exceeded by staged changes,
the application must fail before another mutation. A future transfer that
retains donor serialized state would add the donor policy to the meet. Equal
byte policy differences may not be hidden behind the word “lineage.”

The concrete helper for applying a plan should compute the effective policy,
check `changes.len()` and `replacement_units` against its effective operation
and aggregate bounds, and only then update the edit's effective snapshot policy
and editor retention ceiling:

```text
effective = edit_policy.intersect(plan.planning_receiver_policy)
check already_staged_work <= effective
edit_policy = effective
editor.set_retained_render_limit(effective.max_retained_render_bytes())
```

The checks must precede the policy update so a rejected lower bound leaves the
existing editor and token untouched. The same helper then feeds the ordinary
`replace_text`/nested-object/picture limit checks. A plan's policy is not a
serialized payload; carrying it is what prevents equal-byte reopening from
silently widening the operation's retention or work limits.

`apply_durable` uses wire `PatchLimits`; it starts an edit under the receiving
snapshot's transaction policy and does not create a second DOC retention
policy. `ThreeWayPlan` already intersects the source and both branch
before/after transaction policies. Any future API accepting a patch or
snapshot from a separately opened source with the same bytes must use the
same componentwise meet rather than selecting the after value silently.

`Edit` exposes the retained state without exposing bytes:

```text
retained_render_bytes(&self) -> Option<usize>
release_retained_render(&mut self)
```

The reported value is `Vec::capacity()`, the actual retained allocation
weight charged to the ceiling, not merely serialized length. `release` drops
the allocation, leaves all semantic/package state unchanged, and makes the
next `commit` recompute. A retained length, a zero/`None` result after
overbudget fallback, and release are all observable. If a debug representation
is added, it may report length, capacity, ceiling, generation, and policy, but
never serialized bytes.

The three nested `RevisionEditor::open` assignments in the embedded-object
methods (`add_embedded_object`, `remove_embedded_object`, and
`set_embedded_display_as_icon`) must call the same crate-private opt-in after
reopening. Otherwise a later text edit in one transaction would unexpectedly
lose the caller's retention ceiling. The nested editor's own intermediate
serialized result is not transferred as a token.

## Atomicity, successive edits, and finish behavior

The required state transitions are:

* Before a mutation, the current editor and any token describe one package
  generation. A refusal or candidate validation error leaves both untouched.
* A clone-first candidate begins with no predecessor token. If its final
  common batch succeeds, the package state, generation, and optional new token
  are assigned together. If its rendered capacity is over budget, the
  package state is still assigned and the token is observably empty.
* A second successful mutation starts a candidate with no predecessor token,
  while the original editor keeps its prior token live. Assignment drops that
  prior token and publishes only the second generation's token; a candidate
  failure drops only the candidate and leaves the first token intact. This is
  the successive-edit rule; no prior serialized result may be used for a later
  package state.
* Equal text, equal formatting, equal embedded display, empty batches, and an
  untouched edit do not create a token or render. A public no-op still shares
  the source snapshot allocation and produces the existing no-op patch.
* `release_retained_render` changes only whether final finish recomputes. The
  resulting bytes, snapshot semantics, patch changes, and inverse patch are
  equal to a retained control.
* Nested embedded operations recompute their inspection input from a clone,
  then replace the outer editor with a freshly opened state whose token is
  empty and whose retention ceiling is reapplied.
* `rollback` drops the token. `Edit::commit` consumes it at most once. A final
  strict/public validation or patch error publishes neither a snapshot nor a
  patch; source snapshot identity remains unaffected.

## Exact tests required before the pilot is accepted

### Common editor tests

1. Extend the rendered batch test to prove two existing-stream replacements
   return one exact candidate `Vec`, preserve untouched stream allocations,
   and match the ordinary `finish` bytes and target catalog.
2. Test empty and all-equal batches: `Ok(None)`, no candidate clone, no render,
   unchanged `is_changed`, and exact source bytes.
3. Test malformed, missing, oversized, and package-limit failures. The editor
   and returned handoff must remain absent/unchanged after every failure.
4. Test the optional Data add plus WordDocument/Table replacements as one
   candidate render. If the optional-add variant is not implemented, test the
   explicit missing-Data fallback and assert no retained token is exposed.
5. Compare Reuse copy-through, Reuse layout fallback, and Rewrite control
   outputs for logical package identity, untouched stream bytes, directory
   metadata, CLSIDs, and allocation reuse. A handoff must not alter policy.

### DOC transaction tests

1. `retained_render_bytes` reports the exact retained `Vec::capacity()` after
   a changed paragraph and format operation; `release_retained_render` returns
   `None` and leaves final bytes, semantic text, patch changes, and inverse
   output equal to the retained control.
2. A ceiling of zero and a ceiling one byte below the candidate capacity both
   succeed, report no held render, and produce the same bytes and patch as the
   ordinary route. The limit must not become a refusal.
3. The provisional default/explicit builder and getter round-trip. A direct
   `RevisionEditor::open` reports no retained state, while a body `Edit` uses
   its `Snapshot` transaction ceiling. Reopened embedded-resource editors
   retain the configured ceiling.
4. Exact text, format, icon, empty-edit, and all-equal batch no-ops retain
   exact source allocation and do not report a rendered token.
5. Successive edits prove A token is replaced by B, a failed B edit leaves A,
   release before B is harmless, and the final bytes always describe the last
   package generation.
6. Exercise add/remove embedded object, icon changes, picture insertion and
   picture reversal after an earlier retained text edit. The token must be
   cleared/recomputed and the reopened editor must retain the same bound.
7. Test clone semantics directly: cloning a `RevisionEditor` does not copy or
   share the serialized token, candidate mutation gets only a new token, and a
   failed candidate leaves the original token. `Edit` remains non-Clone and
   `Snapshot::clone` carries no token.
8. Test the sector-policy freshness guard through the common editor seam. A
   policy change after staging must not consume a token made under the old
   policy. The DOC owner currently has no public setter, but the invariant
   must be locked before one is added.
9. Build snapshots with different transaction policies but equal source bytes
   and apply a patch. Assert every transaction field, including retained
   render bytes, is the componentwise minimum and that no bytes are copied or
   re-rendered solely to meet the policies.
10. Prepare edits from separately reopened equal-byte snapshots with different
    transaction policies, join them into one `Composition`, and assert the
    effective policy is the componentwise minimum. Also prepare each transfer
    plan from an equal-byte planning receiver with a zero retained-render
    ceiling, apply it to an equal-byte receiver with an 8 MiB ceiling, and
    assert the effective result retains nothing. Test all four policy members
    and reject a lower operation bound already exceeded by staged changes.
11. Keep the independent final validation test: common candidate validation
    may pass while the final strict/public DOC open is fault-injected to fail;
    no commit result or patch may then be published. The production path must
    not make this failure reachable through a raw byte injection API.

### Candidate evidence

The paired baseline/candidate probe must use the same fresh process schedule,
source bytes, replacements, limits, and owner-drop scopes. It must check exact
output hash, stream/path inventory, directory metadata and CLSIDs, semantic
paragraph/projection witnesses, forward and inverse patches, no-op and failure
controls, successive-edit invalidation, retained capacity, release behavior,
and allocation ownership. Report peak live bytes and the clone-first
old-token-plus-new-render peak separately from boundary-retained bytes. No
phase subtraction, generic cache claim, or production speedup follows from
the candidate route until those gates pass.

## Retention-budget boundary and disposition

The proposed `Vec::capacity()` ceiling meets the retained-state contract: it
is finite, caller-configured through `TransactionLimits`, observable through
`Edit`, releasable without semantic change, uses the already produced common
render allocation, and falls back to recomputation without changing a result,
refusal, or published byte. The custom clone rule prevents a serialized
predecessor from being copied or silently shared through successive edits.

This scalar ceiling does not prove a total process peak bound. The common
renderer and CFB reopen allocate transient buffers, the clone-first candidate
keeps the old token until assignment, and final `Snapshot::open_bounded`
retains a separately owned source/output allocation. Current APIs have no
execution-context budget or hooks that charge all of those live allocations
to one operation. If “retention budget” is interpreted as a proof of total
peak RSS rather than a bound on the caller-retained serialized handoff, that
stronger contract cannot be met by this pilot; the precise obstruction is the
unobservable transient allocation ownership inside common render/reopen and
the independent final reader. The pilot must therefore keep package/resource
limits, allocator evidence, and the explicit old-token/new-render peak gate.
Introducing a global cache, spill file, or unbounded accounting structure
would violate the accepted constraints and is not a remedy.

Proceed with the private DOC-owner handoff and the finite 8 MiB provisional
candidate bound only as an evidence-gated pilot. Keep direct tracked-revision
retention at zero, preserve Reuse and all fallback/no-op behavior, intersect
transaction policies at every meet, and retain the independent final public
validation. No production adoption follows from this design review alone.
