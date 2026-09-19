# 0698 source and lifetime review — borrowed inherited namespace views

`performance_claim: none`

This is the completed source, lifetime, assembly, semantic, and measurement
review of the 0698 candidate against baseline commit
`8f16c57b27fd7878ece718e22db80839583f0098`. The frozen candidate evidence,
follow-up matrix, audit recheck, and repository gates are recorded below. The
final retention decision is rejection because of the reachable refusal-path
regression documented at the end.

## Diff scope

The production diff is confined to the private MCE codec. It changes
`Inherited` from two owning options to two references:

```rust
struct Inherited<'a> {
    ns: Option<&'a Arc<NamespaceLayer>>,
    emitted: Option<&'a Arc<NamespaceLayer>>,
}
```

`hoists` and `for_each_hoisted` now receive those references directly, and
`Inherited::after` uses `cloned()` only for the carried emitted boundary. The
emitted path still clones `ctx.ns.head`, so the child frame continues to own
its namespace scope. `Namespaces::with_local`, the namespace-layer graph,
`Ctx`, `Frame`, `Frame.emitted_ns`, `close`, and the streaming codec are
unchanged.

The old `Inherited` construction was removed from the beginning of `start`.
The candidate constructs a view from the parent frame at each of the three
paths that can create a frame: opaque, direct AlternateContent child, and the
ordinary path. This makes the borrow source the parent frame's pre-local
scope, rather than `c.ns.head`, which can be replaced by local declarations.

## Lifetime proof

The candidate first clones the parent `Ctx`, parses raw attributes, validates
namespace syntax, installs local namespace declarations, expands the element
name, and performs directive and opaque validation. It then computes
`parent_active`. None of those operations needs an `Inherited` view, so
delaying the view does not move an operation that can return an error ahead of
another existing error.

The AlternateContent section now stores its result in `alt`:

```rust
let alt = if let Some(parent) = st.last_mut()
    && let Mode::Alt { .. } = &mut parent.mode
{
    // existing selection, counter, and refusal logic
    Some((active, mode))
} else {
    None
};
```

The mutable borrow of `st` ends when this expression ends. A direct
`AlternateContent` child then enters a separate block and obtains `st.last()`
to construct the borrowed view. The opaque and ordinary paths use the same
block pattern. Each block computes `Inherited::after`, moves the resulting
owned values into a complete `Frame`, and ends before `close(st, frame, ...)`.
Consequently no `Inherited` reference can reach `close`, whose nonempty path
may push into or reallocate the frame vector.

The root path obtains `None` for both fields without manufacturing an owner.
The ordinary path still performs the existing multiple-root check before the
owned frame leaves its block. An error there drops the temporary references
before returning. Empty elements still call `after` before `close`, including
empty `AlternateContent` and branch frames.

The borrow points are semantically the parent's pre-local context:

```rust
let parent = st.last();
let inherited = Inherited {
    ns: parent.and_then(|f| f.ctx.ns.head.as_ref()),
    emitted: parent.and_then(|f| f.emitted_ns.as_ref()),
};
```

The child `c` remains independently owned. Its local declarations therefore
cannot invalidate or alter the view used for inherited hoisting, and the
owned `Frame.emitted_ns` still records either the child's post-local head when
emitted or the parent's nearest emitted boundary when dropped.

## Validation and ordering review

The following source-level ordering is preserved:

- raw attribute decoding and XML/QName syntax checks;
- local namespace syntax, duplicate, binding-limit, and `xmlns=""`
  handling;
- element QName expansion;
- opaque handling and its output write;
- MCE directive parsing, limits, and target validation;
- local directive-layer construction and extension selection;
- AlternateContent attribute validation;
- parent branch selection, choice/fallback counters, selection flags, and
  refusal order;
- compatibility filtering and hoisted declaration emission;
- root tracking and multiple-root refusal; and
- `after` ownership followed by `close`.

The only moved operation is the side-effect-free `Inherited` owner
observation. `after` remains immediately adjacent to frame construction on
each path. The candidate does not alter `c.ns`, directive layers, active
flags, modes, report increments, or output calls. The focused source tests
added for inherited/rebound scopes, opaque/skipped scopes, and refusal
precedence are frozen separately; their passing result is necessary but not
sufficient for this review.

## Evidence assessment

The completed review inspected frozen baseline/candidate assembly and symbols
and confirmed that only the two temporary `Inherited` clone/drop pairs are
removed. The parent `Ctx` owner operations and child-frame `emitted_ns` owner
remain accounted for. Stack reservation and code size are reported without
inferring a broad stack or memory claim.

The completed shared oracle compares exact output bytes and length, `Cow`
ownership, complete `Report`, and typed/debug refusal identity across real
Office XML, deterministic mutations, synthetic scopes, and every
capability/limit profile. The focused MCE suite, refusal matrix, source-bound
namespace/choice/depth/directive limits, declaration-heavy controls, and
repository gates pass. Native timing includes a representative end-to-end
result, tails, and declaration-heavy regression. Allocation counts and bytes
are compared separately; borrowing does not imply an allocation saving.

The semantic oracle, focused tests, refusal identities, owner evidence, and
repository gates satisfy their acceptance conditions. A semantic mismatch,
changed error/report/ownership identity, changed nonempty-frame owner, or
representative regression over the review threshold is a rejection condition;
the measured refusal regression triggers that condition below.

## Frozen ownership and assembly mechanism

The frozen native binaries are bound by
`assembly/baseline.json` and `assembly/candidate.json` and by the more
specific `ownership-assembly/` receipts. Their SHA-256 identities are
`02c34ebf7e912d5a54cfcd18073e108a6a096a77ac575677df6c76aa58b5bc23`
(baseline) and
`2c805a413eb4bb458da46ca61b9ce4283a2797d9712ab97480b23ade9cb7ccf8`
(candidate).

The bounded symbol comparison supports the intended ownership change:

| symbol | baseline | candidate |
| --- | ---: | ---: |
| `start` | 18,083 bytes | 17,833 bytes |
| `drop_in_place<Inherited>` | 118 bytes, present | absent |
| `drop_in_place<Ctx>` | 118 bytes, present | 118 bytes, present |

The baseline `start` assembly contains the parent `Ctx` namespace/directive
clone increments at `0x17805a` and `0x178071`, followed by the temporary
`Inherited` namespace clone at `0x1780a8` and the temporary emitted-boundary
clone at `0x17823b`. Its normal and unwind cleanup paths call the 118-byte
`Inherited` destructor, whose two atomic decrements are retained in
`ownership-assembly/baseline-inherited-drop.stdout`.

The candidate retains the parent `Ctx` clone increments at `0x1777cb` and
`0x1777e2`. The temporary `Inherited` clone sites and destructor call are
absent. The candidate still contains the `Inherited::after` owner clone at
`0x17874b` (the opaque path's `codec.rs:803` mapping) and the corresponding
ordinary-path `after` owner operation around `0x17a2a1`; these operations are
the owned `Frame.emitted_ns`/child-scope owners required by the source proof.
The unchanged 118-byte `Ctx` destructor and its two namespace/directive owner
release paths remain present. The assembly therefore supports removal of the
temporary pair without supporting removal of the parent context or child
frame ownership.

The candidate reserves `0x618` bytes in `start`, versus `0x598` in the
baseline: a 128-byte increase. The `start` function is nonrecursive, so this
is a per-call handler-frame change rather than a per-XML-depth stack bound.
It is a material static tradeoff and must be included in the final decision;
the packet makes no broad stack, RSS, or memory claim. The 250-byte reduction
in the `start` symbol and disappearance of the `Inherited` destructor do not
by themselves predict end-to-end timing. The completed live A/B, allocation,
refusal, and profile records are assessed in the final disposition below.

## Follow-up packet validation review

The frozen `audit.py` now closes the follow-up packet obligations identified in
this review. It requires the driver's fixed raw-output path for every native,
refusal, and oracle timing leg; applies the existing workflow metadata checks
to native outputs; and requires each follow-up refusal case to carry the
expected debug identity with any observed identity matching it. The refusal
metadata must also equal the already validated primary six-leg matrix.

For the declaration oracle, both identity and timing headers must report the
byte length of the retained `mixed.xml`, in addition to matching each other.
The audit also requires the exact packet Python-file census and hashes through
`script-hashes.json`. These checks bind the supporting records to their
intended commands, inputs, semantic metadata, and raw files.

Follow-up validation is therefore closed. These checks establish evidence
admissibility; the measured disposition is recorded below.

## Final retention disposition

The evidence packet is now admissible: the recheck passes the complete audit,
all seven integration gates and six repository evidence gates pass, the
focused MCE run records 96 passing tests on each frozen phase, and the shared
oracle reports zero mismatches across 192 cases, five profiles, and both
binaries. The candidate also removes the temporary `Inherited` destructor,
shrinks `start` by 250 bytes, and improves the representative real one-edit
follow-up medians by 2.45% and 2.30%. The real no-op follow-up improves by
1.26% and 1.12%, and the mixed opaque oracle improves by 1.37% and 1.80%.
All 78 allocation comparisons remain identical, so the speed result does not
come from an allocation reduction.

**Recommendation: reject retention of the current candidate.** The reachable
`early-name-error` refusal is a material counterexample. It mutates an
authored complex package by adding a duplicate XML `name` attribute and
reaches the normal parser error path with the expected and observed
typed/debug identity unchanged. The initial candidate pair regresses its
median by 15.12% (5.33
microseconds), with mean, p95, and p99 also above 12.98%. The independent
300-sample follow-up reproduces a +7.84% median cost (+2.75 microseconds) in
the second ABBA pair, with mean +7.78%, p95 +8.31%, and p99 +6.61%. This is
outside the recorded baseline A/A p50, mean, and p95 noise ranges and exceeds
the packet's +5% review threshold across all four primary statistics. The
first follow-up pair's +0.61% median does not erase the second-pair regression;
other refusal cases do not show the same repeatable pattern.

The over-limit and late-root first-pair spikes do not repeat, and the initial
POI-notes cost falls from +6.40% to +1.74%/+1.46%; those results are retained
as variability and small cost. They do not explain away the duplicate-name
refusal cost. The candidate also carries the measured 128-byte `start` stack
reservation increase, while allocation counts and bytes are unchanged. The
successful real-path improvement is therefore insufficient to accept a
reachable refusal-path regression above the stated threshold. This rejects the
candidate's retention, while leaving the semantic parity and ownership
evidence valid for a future revision.


## Restored-source audit review

Independent review of the rejected-disposition adaptation found no remaining
issues. Candidate and final source maps each bind the exact 602-file census;
the preserved patch reconstructs the candidate, while the working codec equals
baseline and only the three regression tests remain. Candidate integration
receipts bind the frozen candidate bytes. The fresh retained focused receipt
records 96 passes, and the complete 30-driver hash census includes the rejection
transition. `audit-rejected.log` passes. The recommendation remains rejection.
