# Log sections for change 0663

These four blocks are for the coordinator to merge into the shared program
logs. No shared rollup is changed by this branch.

## For `HOTSPOTS.md`

## 0663 — source-backed OLE2 saves retain sector placement and append only when the layout cannot absorb growth

Record: [0663](0663-cfb-sector-layout-policy.md).

Decision 10 of [0652](0652-owner-decisions-for-the-third-wave.md) is now wired
into the source-backed OLE2 route. `OleWriter` has a public `Reuse`/`Rewrite`
policy pair, explicit source adoption, deterministic release-and-allocate
planning and typed fallback reasons. The common OLE editor defaults to Reuse;
DOC's tracked-revision and embedded-object save routes reach it from their
opened source, and the PPT embedded-object finish path adopts its original
bytes explicitly. A newly authored DOC has no source layout and remains the
from-scratch route. The 214-fixture OLE2 corpus parsed 211 fixtures and skipped
3; **209 of 211** reused no-op, same-length and length-changing layouts, and
**195** length-changing cases appended sectors. Every reused output reopened
with every edited stream byte intact, every source directory CLSID retained,
and a clean validation walk. The smoke cases cover both 4095↔5000 MiniFAT
cutoff migrations and shrink/reclaim. Three child processes produced the same
layout digest. A release measurement on `picture.doc` reports Reuse p50s of
255,796–265,081 ns against Rewrite p50s of 241,471–252,786 ns, while paired
A/A floors range from 2.58% to 7.47%; **no speed claim is made**. `kept_sectors`
means placement retention: the generic sink still writes every output sector,
so this record does not claim avoided payload copies or 0617's 13–30% estimate.
`performance_claim: none`. [Record](0663-cfb-sector-layout-policy.md);
[evidence](results/change-0663/README.md).

## For `GOAL_AUDIT.md`

## 0663 — preservation and deterministic fallback decide the OLE2 policy

Record: [0663](0663-cfb-sector-layout-policy.md).

The policy follows `docs/GOAL.md`'s ordering: validate the adopted CFB through
the ordinary parser, preserve source directory metadata and stream identity,
and decline to the established writer whenever a reuse invariant is not proved.
The corpus compares reused directory metadata against the source including
class IDs, compares stream bytes against the edited model and rewrite route,
reopens each result and runs the CFB validation walk. Length growth consumes
released/free sectors before append; both MiniFAT cutoff directions and
shrink/reclaim are tested. Ordered maps, ordered allocation and a fresh-process
digest test cover determinism. The changed `Snapshot::finish` path now carries
the same source layout policy as `ObjectEditor::finish`, closing a route that
would otherwise have silently rebuilt a changed DOC snapshot. The measurement
is retained as a cost boundary: source parsing and planning remain, all output
sectors are emitted through `Write`, and the A/A floor does not separate the
small observed timing difference. No malformed-input defence is weakened and
no partial output is emitted before the reuse plan passes its gates.

## For `REPORT.md`

## 0663 — the ordinary source-backed DOC save now reaches the OLE2 sector policy

Record: [0663](0663-cfb-sector-layout-policy.md).

The implementation closes the route question left by 0617. `litchi-doc`'s
`RevisionEditor` owns a source-backed `litchi-ole-common::object::Editor`; its
ordinary `finish` calls the editor's source-aware render, whose default is
`SectorLayoutPolicy::Reuse`. The embedded-object transaction uses the same
common editor, and PPT's embedded-object writer adopts its original source
before emitting. A changed common-editor snapshot also retains this policy, so
`commit().patch().after()` and `snapshot().finish()` agree on the adopted
layout. The separate public DOC writer authors a new CFB and therefore cannot
reuse source sectors; that limit is documented instead of implying coverage it
does not have. Reopened stream identity, source CLSIDs, cutoff migrations,
append/reclaim and cross-process determinism are all corpus-tested. On
`picture.doc`, Reuse retained 2,795 of 2,828 body sectors for no-op and
same-length edits and appended three sectors after growth; the source and
rewrite byte lengths are recorded beside the timing. These are layout counts,
not avoided payload writes. The release timing has no separable speed result,
so 0617's 13–30% estimate is withheld.

## For `ADR_COMPLIANCE.md`

## 0663 — explicit source adoption keeps ADR 0005 and ADR 0006's preservation boundary

Record: [0663](0663-cfb-sector-layout-policy.md).

ADR 0006's preservation-by-default rule is applied only after explicit source
adoption and a validated directory/geometry match. The planner retains names,
hierarchy, directory entry metadata and class IDs from the source, patches only
the stream allocation fields it owns, and keeps output ordering deterministic.
ADR 0005's retained-state boundary remains explicit: `OleWriter` stores bounded
layout metadata, not an ambient file handle or executor, and the common editor
passes the exact original bytes it captured. A changed directory shape,
geometry, class ID, DIFAT requirement or planner invariant produces a typed
fallback and the existing from-scratch writer. The default policy is a public
breaking 0.0.x API movement accepted by 0652; the `Rewrite` policy remains
available for callers that need canonical compaction. The generic emitter still
writes payload sectors, so `kept_sectors` is not presented as copy or I/O
elision, and no 0617 performance claim is registered. No unsafe code, raw lock,
executor, ambient I/O or weakened resource bound was added.
