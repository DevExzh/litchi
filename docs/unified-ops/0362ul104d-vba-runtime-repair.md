# 0362ul104d — VBA runtime limit and manifest repair

This source-only repair starts from the frozen OOXML runtime candidate
`e130751fc503099c4d5fff74a5746b5099d625e1` (tree
`0251a7e0945eb18d30b80779526ecbe5e0b2fd25`).  The coordinator's common
runtime receipt identified three remaining VBA failures; no compiler, Cargo,
rustdoc, native, or performance command was run in this leaf.

The noncanonical content-types oracle now independently removes the known
plain workbook override and inserts the required macro-enabled main-part and
VBA-project overrides at the manifest close.  This keeps a recipient-only
comment before the generated fragments and still checks every byte exactly.

The structural-limit tests use a tracked Excel source fixture with directory
entries and count the serialized candidates through the existing ZIP reader,
independently of the production retained-metrics planner.  They use exact
candidate budgets and one-under limits that satisfy all `SourceLimits`
cross-field invariants while exercising the intended refusal kinds.  The
replay test consequently distinguishes the member and total-entry ceilings
instead of constructing an impossible policy.

Source inspection, formatting, and repository-size/binary checks passed in this
leaf.  Cargo, compiler, and runtime execution remain intentionally pending for
the coordinator's focused and common runtime rerun.

## Return 0362ul119a

This source-only return starts from the frozen runtime candidate
a8d32513b8e427c6d234ddd17fb836b0e8d2477b (tree
0251ddc8aeef84702604f340406eda671c9b6c54).  The common runtime receipt
reported a relationship-total assertion that compared the complete fixture
graph with only its root .rels member, and a replay fixture that attempted to
publish a topology addition beside opaque directory entries.

The structural assertion now independently sums the root and every present
part relationship token.  The replay test uses a small authored package whose
VBA graph and both custom replacement parts are already source members, then
rewrites it through the existing ZIP reader/writer to add a real custom/
directory entry.  Removing and re-adding known parts, followed by replaying a
known VBA graph update, keeps the targeted preservation route eligible while
still exercising independent ZIP member/entry counts and exact/one-under
archive limits.  No production preservation refusal was weakened.

No Cargo, compiler, rustdoc, native, or performance command was run in this
leaf.  Source inspection, formatting, and repository-size/binary checks were
the only validation performed; the coordinator owns the focused and common
runtime reruns.

## Return 0362ul119b

The first 119a correction still projected a new VBA graph into the POI
directory-bearing source.  That is intentionally refused by the OPC targeted
writer because the directory records are opaque members.  This follow-up
splits the structural limit cases: the authored package reaches
ContentTypeMappings and TotalRelationshipXmlEvents through graph growth, while
a small source-backed package with an existing VBA graph and one real
directory entry reaches ArchiveMembers and ArchiveTotalEntries through an
eligible graph replacement.  Both cases retain independent ZIP counts and
exact/one-under policies; no PreservationUnavailable assertion or production
policy was weakened.

No Cargo, compiler, rustdoc, native, or performance command was run.  The
coordinator owns focused and common runtime reruns from this source-only
commit.
