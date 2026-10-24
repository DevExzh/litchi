# 0362uw — separate ODG joined-publication qualification design

ODG is not part of the four-family shared-triad lane in
`0362uw-odf-shared-triad-qualification.md`. Its returned work is a separate
ordered-composition migration: the historical owner at
`ad52437f5e5e2c71ea476eedd5ddcf61aee2021d` exposes
`Package::apply_joined(JoinedEdits) -> Package`, while the candidate owner
must use one `Edit::join(JoinedEdits)` followed by `Edit::commit()`. The two
endpoints therefore need narrow source-specific probes; compiling one probe
against both APIs would change the baseline lifecycle.

## Required routes

Use a deterministic ODG package assembled by the existing `PackageWriter` and
the same source bytes at both endpoints. Prepare the durable sub-edits outside
the measured closure. Because `JoinedEdits` is consumed, each timed child
rebuilds the bounded plan from cloned/prepared patch inputs; that plan setup is
then identical work on both endpoints.

| route | joined plan | reason |
|---|---|---|
| `package_replace_one` | one structural edit whose durable operation is `package.replace` | isolates the pure replacement path where the new `Edit` may copy source XML before commit |
| `semantic_two_disjoint` | shape-text on one shape plus shape-name on another | verifies independent semantic effects publish in one joined plan |
| `semantic_four_disjoint` | four non-overlapping page/shape property effects | exposes aggregate planning and one-rebuild behavior without substituting a batch API |
| `empty_join` | an empty bounded plan | proves exact no-op/source retention at the new Edit boundary |

The old adapter calls `source.apply_joined(plan)`. The candidate adapter calls
`let mut edit = source.edit(); edit.join(plan)?; edit.commit()?`. No global
marker or generic operation label stands in for a route. Each route supplies
the actual target, value and publication bytes.

## Proof and timing boundary

Before timing, both adapters must prove source immutability, exact output
digest/size, semantic values for every changed target, and deterministic
repetition. The candidate additionally proves the returned Commit diagnostics,
patch applicability and inverse restoration; the old endpoint has no Commit
triad and is not credited with those newer capabilities. A stale-lineage plan
and one conflicting effect are separate untimed refusal checks.

The matched timed child measures two scopes:

1. `join_apply`: prepare the bounded `JoinedEdits` from the already prepared
   route inputs, then run the endpoint's real apply lifecycle;
2. `join_apply_publish`: the same scope followed by the endpoint's output
   publication/readback.

Source opening, deterministic fixture creation, native edits used to prepare
the sub-edit payloads, and digest validation stay outside both clocks. The
runner must retain route-specific output proofs after every ABBA leg. It must
also report source/output byte sizes and the candidate's retained diagnostics
where available; process allocator capture remains an external runner metric.
No route is accepted from a timing ratio until both endpoint proofs and a
quiet-host ABBA run pass. In particular, a faster candidate pure replacement
cannot hide a lost package replacement, and a multi-operation improvement must
still publish all semantic effects from one candidate commit.

## Baseline and candidate handling

The baseline adapter is a temporary external example using only the
`ad52437f5e5e2c71ea476eedd5ddcf61aee2021d` owner API. The candidate adapter
must be copied into the eventual frozen ODG source only after the `Edit::join`
migration and its returned fixes are integrated; its source ref is therefore
intentionally unresolved in this design record. Both adapters consume the same
tracked/deterministic fixture description and `CompositionLimits`, and both
must record their resolved source/lock/probe/feature/toolchain identities.

This design does not claim an ODG result, waive the returned migration, or
modify the original W8 f2 packet. It is a separate follow-up lane to run after
the ODG source correction reaches candidate freeze.
