# 0730 DOC validated-render handoff implementation review

Status: **scoped bounded-pilot adoption recommended; all packet gates pass**
(2026-09-22). This is a read-only review of the candidate in `litchi-doc` and
`litchi-ole-common`; no Cargo or native command was run and no production file
was changed by this review. The focused integration file
[`doc_render_handoff.rs`](/home/zhuhe/code/litchi/crates/litchi-doc/tests/doc_render_handoff.rs)
now covers policy intersections, retention boundaries, release, successive
edits, no-ops, refusal, patches, composition, transfer, and representative
reopens. The retention allocation matrix and fixed ordinary A/B are complete;
the inherited oracle, independent audit, all 20 corruption controls, and
offline terminal replays pass after cleanup. The recommendation below is
limited to the measured DOC owner scope; the root report carries the final
packet presentation.

The public policy documentation is now explicit: `TransactionLimits::default()`
enables the 8 MiB pilot retention ceiling, while `TransactionLimits::new(...)`
starts with retention disabled until the caller selects a ceiling. The focused
oracle helpers explicitly override the default to zero, including the
same-sequence multi-replacement reference, so those comparisons cannot consume
a retained candidate accidentally.

Review chronology is retained because the candidate was edited during review.
The initial pass occurred before `doc_render_handoff.rs` was added and reported
the missing owner tests. An intermediate source snapshot briefly attempted to
read `transaction_limits` from a shadowed embedded-object donor; that finding
is superseded by the current receiver-policy implementation below and is not a
live compile finding. This refreshed review is against the following current
working-tree SHA-256 files: `body_text.rs`
`c1447266d7a4e76c8cbc49d9b7bd585737ff4532b3156fb977452f908261b18a`,
`package.rs`
`ee9672af957a205f100cd0ca4211321ea8042c07520bbe16d0ea1a6aa916999a`,
`object/editor.rs`
`4f01d8b11b79c18db1a0a750f96d010b7041dbd2062bbbe83c601d3ab860814f`, and
`doc_render_handoff.rs`
`797c822aae89f60b69c1f05c3609a60b31d80a5ef859e81eceff6c68b9cadc3d`.

The candidate has the right basic ownership shape. `RetainedRender` stores one
`Option<Vec<u8>>`, its custom `Clone` is empty, and `RevisionEditor::finish`
consumes that exact allocation. The mutation entry points in
`tracked_revision/package.rs` are clone-first, and the publication boundary
clears the candidate token before installing a replacement. `Snapshot::editor`
and the resource-owner `Edit::reopen_editor` both apply the transaction
ceiling. This keeps a rendered handoff attached to the owning editor state,
rather than to a byte-equal `Lineage`, and avoids cloning a rendered result.

## Source atomicity resolution

The intermediate post-publication error is now addressed. The clone-side
`RevisionEditor::replace_with_picture_graph` computes the installed graph and
returns `Ok(None)` before `candidate.commit()` when a supplied graph with a
`data_offset` does not match. `Edit::install_picture` maps that `None` to the
existing `Error::Conflict` before assigning its candidate editor. This keeps
the old semantic state and retained-capacity report on that failure without an
extra clone. The body-owned
`retained_picture_mismatch_is_atomic_and_missing_data_handoff_is_current`
test now exercises that internal guard for both inline and floating pictures,
and checks token, change-count, and serialized bytes before and after the
refusal. Its valid path also checks the missing-`Data` handoff against a
zero-ceiling recomputation.

The package clone/finish test verifies that cloning drops the retained token
and that the original `finish` moves its exact allocation. The resource-reopen
test verifies clearing and re-enabling retention, and the diagnostics test
keeps the independent public-reader validation at final commit. These close
the previously outstanding source-level controls.

The current transfer implementation records only the receiving snapshot's
`TransactionLimits` in each plan (around
[`body_text.rs:1097`](/home/zhuhe/code/litchi/crates/litchi-doc/src/body_text.rs:1097),
[`body_text.rs:1143`](/home/zhuhe/code/litchi/crates/litchi-doc/src/body_text.rs:1143),
and [`body_text.rs:1205`](/home/zhuhe/code/litchi/crates/litchi-doc/src/body_text.rs:1205)). Applying the plan intersects that value with the receiving edit. This
is the receiver-owner policy: a donor is read-only and its limit does not
authorize how the current receiver edit retains output. A byte-equal receiver
with a stricter policy is still reduced at application, and the focused test
exercises that case. No donor-policy intersection is required.

## Policy propagation and byte-equal lineage

The candidate now carries the owner policy through the paths that can meet
byte-equal snapshots:

| Path | Current policy result | Review result |
| --- | --- | --- |
| `Patch::apply` | `min(source, before, after)` | Correct shape; test low/high equal-byte snapshots. |
| `ThreeWayPlan` | `min(source, both patch endpoints)` | Correct shape; test the retained ceiling as well as operation limits. |
| `PreparedEdit`/`Composition` | Policy is stored in each prepared edit and intersected on successful joins; commit applies the aggregate to a fresh source edit. | Correct shape after the current edit; test rejected joins leave the aggregate unchanged. |
| Text, embedded-object, and picture transfer plans | The plan records the receiver policy; application intersects it with the receiving edit. | Correct under the receiver-owner policy; the equal-byte strict-plan/broad-receiver test covers the boundary. |

`Lineage(Arc<[u8]>)` compares bytes, so allocation identity cannot be used as
the policy boundary. The retained `Vec` must stay in `Edit`/`RevisionEditor`;
none of the plans may acquire a render token. The new policy fields follow that
rule. The tests must prove that a high-ceiling donor cannot make a low-ceiling
receiver retain output, and that a low-ceiling prepared edit/transfer cannot be
silently widened by a high-ceiling byte-equal composition or receiver. The
composition case is covered by the new aggregate field, and the transfer case
is intentionally receiver-owned and covered by the focused test.

## Retention and lifecycle findings

The finite policy uses `Vec::capacity()`, filters a successful common handoff
against the ceiling, exposes the held capacity, and has explicit release. An
over-ceiling result falls through `RevisionEditor::finish` to the existing
package render, so it does not refuse supported work. The common batch API
returns the final validated render after candidate publication and retains
untouched stream `Arc`s. The final DOC commit still performs strict-owner and
public-reader validation when the bytes differ from the source.

The earlier `data_stream_added` suppression has also been removed. The common
returned bytes are the final render of the isolated candidate, including a
newly created `Data` stream, so retaining that `Vec` is safe and does not need
an optional common `add` API or a blanket fallback.

The source audit found the intended invalidation shape:

- ordinary text, formatting, picture, and revision mutations clone the editor;
- clone drops the old token without copying its `Vec`;
- resource-owner operations render a clone, then reopen through the helper that
  reapplies the current ceiling;
- true body no-ops return before mutation; and
- rollback/drop naturally releases the owning editor and its `Vec`.

The remaining memory bound is a retained-state bound, not a process-peak bound.
On a successful second mutation, the old `Edit` token remains live while its
clone renders the new candidate; assignment then drops the old token. Thus two
render allocations can overlap transiently even though only one token is
stable afterward. The multi-operation allocation lane must measure this and
must not claim that the 8 MiB ceiling bounds peak live memory. A failed
clone-first mutation should retain the old token; a successful mutation must
report only the newest state. Resource-owner reopen, explicit release, and
over-ceiling recomputation need the same checks.

## Retention-analysis evidence

The separate retention probe completed all 24 processes once: ordinary and
allocator lanes, both DOC cases, one and two edits, and zero/default/release
retention. Every capture returned successfully, matched its zero-retention
reference output byte-for-byte, and obeyed the expected capacity and release
state transitions. The first checker invocation rejected the completed packet
because its CLI labels (`default`/`release`) differed from the report labels
(`default_8mib`/`release_8mib`). That failed checker and explanation are
archived in [`retention-validator-attempt-0.md`](/home/zhuhe/code/litchi/docs/performance/results/change-0730/retention-validator-attempt-0.md);
the corrected analysis consumed the existing captures without rerunning or
modifying them. The current receipts are
[`retention-analysis.json`](/home/zhuhe/code/litchi/docs/performance/results/change-0730/retention-analysis.json)
(`88b30088a674be462b132fdae10d7cd4768a206a9caa298d93fb03bf170bbad3`) and
[`retention-captures/manifest.json`](/home/zhuhe/code/litchi/docs/performance/results/change-0730/retention-captures/manifest.json)
(`e12bedd4dbb59b26328093d11632254548413b32c908290e6d709c5f137c14c3`).
The surrounding custody receipts bind this analysis to retention build 2 and
the final candidate build 5; the qualification receipt is `pass`, and the
quality-3 manifest contains 13 zero-exit gates. Those commands were not rerun
as part of this read-only review.

The allocator region reports peak live bytes relative to live bytes at entry;
it is not RSS. Its `retained_bytes` field is the live-byte delta at region exit,
not the `RetainedRender` capacity. The exact two-edit default-versus-zero
results are:

| Case | Zero peak live | Default peak live | Increase | Default handoff capacity | Boundary retained bytes (zero/default) |
| --- | ---: | ---: | ---: | ---: | ---: |
| `FloatingPictures.doc` | 3,182,172 | 3,714,077 | +531,905 (+16.715%) | 587,776 | 353,280 / 353,280 |
| `NoHeadFoot.doc` | 236,249 | 273,113 | +36,864 (+15.604%) | 36,864 | 28,672 / 28,672 |

Rounded to two decimal places, the default two-edit peak increases are
**+16.72%** for `FloatingPictures.doc` and **+15.60%** for `NoHeadFoot.doc`.

The one-edit default peak is unchanged from zero for both cases. The explicit
release route matches zero for every recorded allocation field, and the
two-edit default still reduces total allocated bytes and calls versus zero
(`FloatingPictures`: 17,387,098 bytes and 17,765 calls versus 18,795,707 and
18,380; `NoHeadFoot`: 1,217,121 and 2,198 versus 1,343,516 and 2,447).
Those reductions are raw observations, not a speedup claim. The stable
boundary retained bytes are identical because every route returns the same
committed output ownership; the flagged cost is the transient stage-2 peak.

The two-edit peak increase is the expected old-token/new-candidate overlap:
stage 2 begins with the stage-1 handoff, renders a candidate, and only then
replaces the owner token. It is finite in this ownership model because one
successful mutation drops the prior token, so successive edits do not retain
one rendered allocation per edit. The result is an explicit greater-than-5%
peak review flag under the packet hypothesis, not a correctness or atomicity
failure. I consider it acceptable for the bounded pilot only with the scope
made explicit: the 8 MiB ceiling bounds one caller-retained serialized
handoff, while transient renderer/reopen allocations and the old-token/new
candidate overlap remain outside that ceiling; this does not promise a total
process-memory or RSS bound. Keep zero retention for direct revision-editor
use, preserve explicit release and recomputation fallback, and do not generalize
these two fixtures to arbitrary producers. The completed ordinary A/B supports
the bounded-pilot recommendation, and any caller that later requires a total
peak-memory budget should select zero retention rather than silently accepting
the peak cost. The flag remains visible.

## Main A/B evidence and adoption recommendation

The fixed main packet completed 44 main capture processes plus four before
controls (48 total): 36 native timing processes and eight allocation
processes in the main packet. The final
[`analysis.json`](/home/zhuhe/code/litchi/docs/performance/results/change-0730/analysis.json)
(`ee63cbc30cf7721798ffc958f8d3f1922da6ece0694a469e4a7e47c560c13fe5`) and
[`captures/manifest.json`](/home/zhuhe/code/litchi/docs/performance/results/change-0730/captures/manifest.json)
(`61f8a4981a6f0db20517da59c806748eb91aa496cdbb928db59b00acfe48d25c`)
record exact output identity, custody, and the inherited DOC oracle; the
independent audit also passes.

The central timing result is consistent on `NoHeadFoot.doc`: all six AB/BA
pairs across three cycles are faster at p50, with improvements from 11.890%
through 13.876%. `FloatingPictures.doc` is mixed, ranging from a 14.860%
improvement to a 1.272% regression at p50; its corresponding +0.852% mean in
the positive pair remains below the 5% review threshold. Every candidate AB/BA
pair has no positive p50 or mean regression above 5%. All AA p50 and mean
controls stay within 5%.

The complete flag inventory is retained here so the favorable central result
does not hide tail or allocator variation:

| Evidence | Flags | Interpretation |
| --- | --- | --- |
| AA timing controls | `docfloat` cycle 2: `p99`, `maximum`; `docnohf` cycle 0: `p95`, `p99`, `maximum`; cycles 1 and 2: `p99`, `maximum` (cycle 2 also `p95`) | Baseline-control tail variation; all AA p50/mean values remain within 5%. |
| Candidate timing | `docfloat` AB cycles 0–2: all five fields; BA cycles 1–2: all five; BA cycle 0: none. `docnohf` AB cycle 0: `p50`, `mean`, `p95`; AB cycles 1–2 and all BA cycles: all five fields. | Every flagged candidate timing delta is negative, meaning faster; no flagged positive central regression exists. |
| Main allocation A/B | `docfloat`: `allocated_bytes`, `deallocated_bytes`; `docnohf`: those two plus `allocation_calls` | Flags are reductions: float allocated/deallocated bytes −9.009%/−9.213% and calls −3.903%; NoHeadFoot allocated/deallocated bytes −11.796%/−12.111% and calls −13.057%. `peak_live_bytes` and `retained_bytes` are unchanged. |
| Retention two-edit default | Both cases flag peak live: +16.72% float and +15.60% NoHeadFoot; one-edit default and explicit release have no peak flag | Expected old-token/new-candidate overlap, already reviewed as a bounded retained-state cost. |

I recommend scoped adoption of the validated-render handoff. The
recommendation is for the public `litchi-doc::body_text` owner route with the
finite 8 MiB default ceiling, one caller-retained validated `Vec`, capacity
accounting, explicit release, zero/over-capacity recomputation fallback,
strict-owner and public-reader validation, and the existing CFB `Reuse` policy.
Direct `RevisionEditor` use remains at zero retention. The implementation is
not a general CFB cache, a total process-memory budget, or an authorization for
donor transaction policy; transfer policy remains receiver-owned.

The performance claim is limited to this packet's two DOC fixtures and public
lifecycle. `NoHeadFoot.doc` supplies the consistent p50 improvement;
`FloatingPictures.doc` supplies a mixed result with no greater-than-5% central
regression.
The 16.72%/15.60% two-edit allocator peaks remain an explicit limitation: they
are acceptable under the stated retained-state scope, but any caller that
requires a total peak-memory or RSS ceiling should select zero retention. No
broad producer, format, or speedup claim follows from these captures.

## Packet closure

The focused tests, retention matrix, exact main oracle, and independent audit
establish the source, ownership, and measured two-fixture behavior. Root owns
the final report presentation; cleanup, corruption controls, and offline
terminal replays are complete. No additional implementation or measurement
blocker was found in this refresh.
