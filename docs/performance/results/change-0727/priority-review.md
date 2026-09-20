# 0727 non-iWork priority review

This is the authoritative read-only queue review for the next bounded
work-elimination measurement. It uses the optimization order and preservation
rules in [GOAL](../../../GOAL.md), the current [hotspot inventory](../../HOTSPOTS.md),
and the remaining CRUD coverage in [CRUD_COVERAGE](../../CRUD_COVERAGE.md).
No new timing is taken here. The 0723–0726 XLS checkpoint and empty-slot
variants remain rejected; 0727's native tail replication is diagnostic and does
not reopen that queue.

The first draft is preserved as [priority-review-initial.md](priority-review-initial.md).
It incorrectly proposed the XLSX shared-string admission after 0667 had
already landed it; this file removes that proposal and rechecks the surviving
opportunities against changes 0667–0726.

## Ranked queue

### 1. Qualify a length-changing OLE2 copy-through writer for DOC/PPT

This is the leading hypothesis for fresh cross-format qualification; the old
measurements do not establish its current payoff.
The current source-backed CFB overlay and the later 0663 sector-layout policy
cover same-length overlays or a layout-aware rebuild. 0663 explicitly leaves
length-changing edits on the planned-artifact route. Change 0617's native
baseline makes the remaining gap format-selective: `FloatingPictures.doc` has
**81.4% unchanged stream bytes**, the container rebuild is **13.25% of commit
cycles**, and work proportional to untouched streams is **12.69%**. The
smaller `NoHeadFoot.doc` case is **29.99%** container and **26.04%**
untouched-stream work. `45543.ppt` has 16.9% unchanged bytes and a 7.26%
container share. XLS is explicitly a poor target: 4.8% unchanged bytes and a
1.07% container share, inside that measurement's native floor. The rebuild also
peaks at 5.2–5.9× the artifact size, so the opportunity includes retention
pressure as well as copying.

The concrete source seams are:

* The existing from-scratch writer is
  `crates/litchi-cfb/src/writer/core.rs:291` and `:1122`; the sequential writer
  is in `writer/sequential.rs:614` and `:757`.
* The common length-changing route enters at
  `crates/litchi-ole-common/src/object/editor.rs:80`, replaces streams at
  `:315`/`:372`, and rebuilds in `:657`. `object/codec.rs:453-504` contains
  `render_with_layout` and the already-landed **same-length**
  `render_copy_through`; that function is not the proposed length-changing
  path.
* The reusable validated primitives are
  `crates/litchi-cfb/src/overlay.rs:569` (`ValidatedOverlayPlan`), `:583`
  (`ComposedOverlaySource`), `:733` (same-length planner), and `:1118`
  (`finish_overlay_plan_with_owner`), with the source-backed owner in
  `crates/litchi-ole-common/src/source_backed_overlay.rs:22`.

The first action is a current-source DOC/PPT baseline and memory/work attribution
under the implemented 0663 Reuse policy, followed by source qualification and
a bounded pilot only if the measured opportunity remains material. The candidate must be a separate fourth
publication path over a validated `SharedOleFile`: existing stream replacement
only; no create/delete/move/rename or storage topology change; bounded
replacement/output/readback budgets; and initial refusal for DIFAT, signed,
encrypted or DRM-marked sources. It should compose unchanged source bytes with
changed stream spans and an appended sector tail, copy directory records
instead of rebuilding them, reopen before publication, and verify every
untouched stream byte-for-byte. Existing fingerprint brackets, source-version
checks, FAT/MiniFAT/chain validation, typed errors and atomic publication stay
in force. The physical placement question left open in 0617 was subsequently resolved
by owner decision 0652 item 10 and implemented in 0663: preserve the current
Reuse/Rewrite policy, deterministic release-and-allocate behavior and fallback
contract. Do not request that already-settled decision again. The later 0663
picture.doc Reuse measurement regresses about 40% against Rewrite, so the old
0617 fractions cannot serve as a current matched baseline. The pilot must
report DOC/PPT native work and peak live bytes separately; it must not promote
the falsified XLS case into this route.

This fills the CRUD gap where the matrix currently covers same-length OLE2
stream edits and metadata moves but leaves broad length-changing DOC/PPT
producer saves partial. It is also named directly in GOAL's legacy CFB work and
definition of done.

Evidence: [0617 design and native baseline](../../0617-cfb-copy-through-writer-design.md),
[0663 sector-layout policy](../../0663-cfb-sector-layout-policy.md),
[GOAL's CFB copy-through item](../../../GOAL.md), and the [OLE2 CRUD row](../../CRUD_COVERAGE.md).

### 2. Qualify a bounded replay proof for the remaining PPTX cross-slide plan work

Change 0656 removed the second candidate serialization/deflate under a bounded
retained-archive limit. It deliberately left the in-memory candidate graph and
the complete replan intact. On the media-rich lifecycle, planning remains
about **297–301 ms p50** while apply is about **89–94 ms p50**; archive reuse
removed roughly 23–27% of whole native lifecycle time in the accepted windows.
The next measured term is therefore candidate construction and replan work,
rather than another archive-retention attempt. Later 0662 changed bounded
publication deflate behavior but did not remove this planner work; 0703–0704
concerned opened-PPTX capture/projection reuse, not this cross-copy planner.
The cross-copy path remains partial in CRUD coverage outside its bounded
dependency closure, so this has narrower reach than the OLE2 opportunity.

The source path is explicit:

* `crates/litchi-pptx/src/opened/cross_copy_plan.rs:565` plans and retains a
  bounded archive; `:596` applies the plan after semantic and physical revision
  checks.
* `:830` validates the application candidate, `:907` performs the complete
  `prepare_cross_slide_copy_for_slides` refusal/closure work, and `:1157`
  rebuilds the candidate graph and reopens the serialized package.
* `crates/litchi-pptx/src/opened/patch.rs:848` is the final
  `validate_candidate` proof; `cross_copy_plan.rs:2151` owns the physical
  package fingerprint and `:2184` its snapshot cache.

The next measurement should first split `build_candidate` into graph/closure,
serialization and reopen/capture costs, then test whether a compact,
deterministic replay token can remove only graph reconstruction on the clean
owned-source route. It may use a retained archive only when the existing
`max_retained_candidate_bytes` bound admits it and all source/destination
semantic and physical revisions still match. It must retain the 0646/0656
constraint that a plan does not retain a `Snapshot`, and it must not bypass
signature, macro, unknown-member, MCE, protection, collision, limit,
`validate_before`, `validate_after`, target-revision or typed-refusal checks.
Dirty/foreign ingress, missing retention, changed limits, or any proof mismatch
must fall back to today's full replan. If those checks cannot be preserved with
a replay token, close this candidate rather than trusting the archive or adding
a hidden cache; a durable patch or public snapshot representation is out of
scope.

Evidence: [0656 remaining planning cost](../../0656-pptx-cross-copy-candidate-budget.md),
[0646 proof boundary](../../0646-pptx-cross-copy-candidate-retention-design.md),
[0662 bounded publication deflate](../../0662-parallel-changed-member-deflate.md),
and the [PPTX copy-closure CRUD row](../../CRUD_COVERAGE.md).

### 3. Re-qualify a source-preserving DOCX `alt::scan` ownership seam

This is a lower-priority, explicitly gated qualification after the two
container/candidate opportunities. Change 0711's event/resolver borrowing
pilot was rejected and production was restored: generated edit p50 improved
8.15–8.86%, but the admitted `NumberedList` pair improved only 2.888%, missing
the 3% gate; one lifecycle p99 rose 18.64%, and allocation reductions were
diagnostic only. The evidence still identifies an attributable parser boundary:
historical `alt::scan` work was 37.33% of the generated document-body owner and
18.18% on the contrasting admitted real fixture. This can remove event and
resolver ownership work if a narrower source proof is found, but the rejected
pilot itself is not a candidate to rerun unchanged.

The source entrypoint is `crates/litchi-docx/src/alt/codec.rs:107`, with the
writer-local caller and differential seam in
`crates/litchi-docx/src/writer/doc/package.rs:17` and the retained 0722
fusion witnesses. Later 0713 changed the shared MCE substring helper, while
0715/0721 rejected broader DOCX fusion and 0722 retained only a writer-local
fusion; none of those changes accepts 0711's ownership patch.

The next step is a source-bound design/measurement only: prove a no-`altChunk`
or exact event-context route that keeps source ranges, namespace choice,
unknown markup, relationship validation, refusal order, limits and allocation
failure behavior unchanged. A candidate may borrow the resolver only within
that proven context and must fall back to the current owned scan otherwise.
Repeat generated and `NumberedList` native edit/lifecycle p50, mean, p95/p99,
allocation, differential and exact-output gates before any production change;
the 0711 p99 and near-miss gate remain hard review inputs. Do not broaden this
into structural DOCX scan fusion.

Evidence: [0711 rejected pilot](../../change-0711.md), [0722 retained writer-local
fusion](../../0722-docx-writer-local-fusion-pilot.md), and the [DOCX CRUD rows](../../CRUD_COVERAGE.md).

## Later-change screen and exclusions

The shared-string proposal from the initial draft is closed: 0667 implemented
stable-index admission, one retained internal table, lazy materialization and
the 64 MiB bound. Current source confirms `SharedStringsState` capture and
table publication, so 0667 is evidence of a completed seam rather than a next
opportunity. Its after-only producer timing is not promoted to a speed claim;
any real-producer follow-up is coverage/measurement work, not a new admission
design.

The rest of the later queue was screened as follows: 0668 packed-cell and SST
locator scratch work landed; 0669 removed the XLSB candidate reparse; 0672
landed stored-cell reservation; 0701 and 0713 landed exact MCE search changes;
0702 retained borrowed MCE views but left tail flags; 0704 retained bounded
slide transforms; 0706 XML-auditor bookkeeping and 0708 validator-name storage
failed native gates; 0710 fixed custom-properties preservation; 0711's
`alt::scan` pilot failed as described above; 0715 and 0721 rejected broader
DOCX fusions; 0722 retained the writer-local fusion; and 0723–0726 rejected
the XLS checkpoint variants. None of those dispositions changes the three
bounded opportunities ranked here.

Also excluded are the already-landed XML namespace writer/fragment work (0653),
selected-cell early stopping (0658), PPTX revision memoization (0655), PPTX
candidate archive retention (0656), and CFB Reuse/Rewrite sector policy (0663).
The MCE `Namespaces::with_local` empty-vector clone remains a reserve
experiment only: 0695/0697 attribute it, while 0698 rejected the broader
borrowed-view pilot and 0702 retains unresolved tail evidence.
