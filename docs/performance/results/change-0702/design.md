# Change 0702 design — borrow inherited MCE namespace views on the marker-search baseline

Status: static design record captured before candidate measurement. This
document does not claim a speedup, semantic equivalence, or production
retention. The candidate source and the final disposition belong in
[`code-review.md`](code-review.md) after the frozen source, assembly, oracle,
and timing evidence has been reviewed.

## Target and reason for this retry

The current production source is the retained 0701 MCE marker precheck at
revision `72d6f4500d906b9a5d6d4a0e571b0bb840ab1e68`. Its early predicate uses
the private outlined `contains_mce_namespace` helper: `memchr` finds only
possible first-byte positions and `starts_with` checks the fixed URI. The
candidate in this packet does not change that helper or its call site.

This packet re-measures the narrow 0698 namespace-ownership hypothesis against
that newer baseline. The old 0698 candidate removed two temporary `Arc` clone
and drop pairs for the per-element `Inherited` view, but it was rejected after
a reachable `early-name-error` refusal regressed beyond the review threshold.
The new baseline matters because 0701 changed the marker dispatch and the
generated layout of the surrounding processor. The old timings and assembly
cannot be transferred to this source state. The retry is therefore a fresh
experiment, not a claim that the rejected 0698 candidate becomes acceptable.

The intended candidate changes only the temporary `Inherited` view from
owning options to references borrowed from the parent frame. The parent
`Ctx` clone, child `Frame` ownership, namespace-layer graph, marker helper,
streaming path, and all public boundaries remain unchanged.

## Smallest intended source change

The baseline currently has:

```rust
struct Inherited {
    ns: Option<Arc<NamespaceLayer>>,
    emitted: Option<Arc<NamespaceLayer>>,
}
```

The candidate may change only this temporary representation and its three
construction sites to the equivalent borrowed form:

```rust
struct Inherited<'a> {
    ns: Option<&'a Arc<NamespaceLayer>>,
    emitted: Option<&'a Arc<NamespaceLayer>>,
}
```

`hoists` may pass those references directly. `Inherited::after` must still
return an owned `Option<Arc<NamespaceLayer>>`: an emitted child clones
`ctx.ns.head`, while a dropped or unwrapped child clones the nearest emitted
boundary with `cloned()`. The optimization removes only the temporary view's
owner operations. It must not remove the parent `Ctx` clone or the owned
`Frame.emitted_ns` result.

The source diff must remain confined to
`crates/litchi-ooxml-common/src/mce/codec.rs`. The existing 104-test focused
source is the retained semantic witness and is unchanged in this experiment;
no new test behavior is planned. There must be no Cargo or lockfile change,
dependency addition, public item, unsafe code, global state, cache, parser or
frame rewrite, streaming-codec change, limit movement, API change, or change
to the 0701 marker helper. `find_bytes`, active-offset processing, and the
marker-free borrowed fast path remain out of scope.

## Lifetime and ordering proof obligation

The borrow must come from the parent frame's pre-local namespace scope, not
from the cloned child context. The child context can replace `c.ns` after
raw namespace declarations are validated; borrowing from `c.ns.head` would
describe the wrong scope and would prevent that replacement. The required
ordering is:

1. Decode attributes and perform the existing XML and attribute checks.
2. Clone the parent `Ctx`, collect local declarations, install them, expand
   the element name, and perform every QName, directive, opaque, and limit
   check in its existing order.
3. Finish every mutable borrow of `st`, including AlternateContent choice,
   fallback, counter, and selection updates.
4. Construct the borrowed `Inherited` view from `st.last()` only after those
   mutable borrows end. At the root, both fields remain `None`.
5. Run the existing `write_start` call, call `Inherited::after`, and move the
   resulting owned values into a complete `Frame`.
6. End the borrow block before calling `close(st, frame, ...)`. No reference,
   alias, closure capture, or helper return derived from `Inherited` may reach
   `close`, which may reserve or push into the frame vector.

The opaque path, direct AlternateContent-child path, and ordinary path each
need this owned-frame-before-`close` boundary. Empty elements must evaluate
`after` before `close`, including empty branch frames. A source rewrite that
depends on non-lexical lifetime inference while leaving the boundary implicit
does not satisfy this review.

## Semantic invariants

For every byte input, capability profile, limit profile, and parser outcome,
the following must remain unchanged:

- the inherited namespace head is the parent's pre-local head;
- local declarations still install a parent-linked layer, reject duplicate or
  invalid declarations, enforce the namespace-binding bound, and preserve
  `xmlns=""` and implicit `xml` accounting;
- `hoists` and `for_each_hoisted` see the same owner identities and nearest
  emitted boundary;
- an emitted child owns its post-local namespace head in `Frame.emitted_ns`;
  skipped, unwrapped, opaque, AlternateContent, Choice, and Fallback frames
  retain the nearest already-emitted boundary;
- QName expansion, directive validation, opaque preservation, branch
  selection, error precedence, output bytes, complete `Report`, and `Cow`
  ownership are unchanged; and
- the 0701 marker precheck remains byte-oriented, bounded, and lexically
  equivalent to the previous window predicate, including malformed bytes,
  comments, text, final complete windows, and limit ordering.

Borrowing the temporary view does not make the codec or its output state
retained. It retains no pointer after the call, no cache entry, source owner,
lock, execution context, or worker. Existing source, output, depth,
namespace, directive, and allocation limits remain authoritative.

## Existing focused tests are sufficient for this source scope

The 104-test `mce::` suite already contains the three 0698 regression tests:

- `inherited_scope_survives_selected_alt_and_unwrapped_branch_rebindings`
  covers selected AlternateContent, skipped content, unwrapped content,
  prefix rebinding, and exact output/report values;
- `opaque_and_skipped_scopes_keep_the_nearest_emitted_namespace` covers
  opaque descendants, nested rebinding, skipped wrappers, and the nearest
  emitted boundary; and
- `nested_context_errors_keep_first_refusal_precedence` covers QName,
  AlternateContent-order, and directive-related refusal precedence.

The surrounding focused suite also covers empty namespace scopes, exact
namespace limits, empty elements, inherited directives, preservation and
ProcessContent, dropped wrappers, selected and rejected branches, malformed
input, and active-offset use of the MCE processor. The three 0701 marker
dispatch tests cover arbitrary-byte lexical boundaries independently. Since
0702 changes no semantic test source and the candidate's only behavior is
owner representation, adding another source test would duplicate these
existing witnesses rather than close a specific untested invariant. The
oracle remains responsible for exact output, ownership, report, and typed
error parity across its full corpus.

## Cost model and risks

The candidate removes two temporary `Arc` clone increments and the matching
drop operations per processed element view. This is an owner-operation
reduction, not an allocation or RSS guarantee: `Arc::clone` need not allocate,
and the parent context plus child-frame owner remain. The candidate can still
change generated code, register pressure, stack reservation, code placement,
and branch behavior. In particular, the 0701 marker helper already changed
the enclosing processor layout, so 0698 assembly and timing numbers are not
valid controls for this retry.

The principal risk remains refusal and marker-positive paths that construct
frames only briefly. The rejected 0698 packet recorded a repeatable
`early-name-error` regression in that class. The retry must measure that case
again, along with late errors, MCE-positive refusals, declaration-heavy
controls, and real edit/no-op/two-target workflows. Marker-free inputs that
return before the parser should have the same logical path, but code layout
effects still need to be observed rather than assumed away.

No end-to-end gain, allocation reduction, stack saving, RSS reduction, or
universal MCE/Office claim may be inferred from the source diff or from the
disappearance of a destructor symbol. A call-frame or symbol-size change is a
diagnostic result and must be reported separately from peak process stack or
memory bounds.

## Required evidence before retention

The packet must bind a fresh baseline and candidate source census, build
inputs, lockfiles, corpus and environment, and binary identities. The
baseline is the complete 0701 production source with the 104-test source
unchanged; the candidate is the same source plus only the borrowed `Inherited`
codec diff. Both focused receipts must pass 104 tests with warnings denied.

The required semantic and resource evidence is:

1. The shared 192-input × 5-profile oracle must report exact output bytes and
   lengths, `Cow` ownership, complete `Report`, and typed/debug error identity
   parity for both binaries.
2. The 0701 marker-control family, 13-workflow native matrix, ten-case refusal
   matrix, independent allocations, profile counters, native size, processor
   assembly, and stack diagnostics must all be freshly rerun. The marker
   helper's setup and long-scan controls remain part of this batch.
3. All seven integration and six evidence gates must pass on the frozen
   candidate, with no Cargo overlap or partial gate treated as completion.
4. Only after those gates are terminal, the independent six-case native and
   ten-case refusal follow-up must run with its four ABBA legs, 300 samples,
   and ten warmups. Its audit must recompute raw statistics, identities,
   commands, sample counts, and pair deltas.
5. The final review must disclose every primary-statistic trigger around the
   goal's approximately 5% threshold, the prior early-name result, all
   declaration-heavy costs, code/stack changes, and any baseline drift. A
   geometric mean or one favorable leg cannot erase a trigger.

## Retention and rejection rule

Retention is conditional on exact semantic parity and a statistically
credible, practically useful representative benefit with no unexplained
resource or refusal cost. The goal's approximately greater-than-5% latency or
throughput regression and 5% peak-RSS review triggers apply to each named
scenario and metric. A candidate with no material end-to-end gain is not kept
for an owner-operation saving alone.

The candidate must be rejected if the borrowed view repeats the reachable
early-name/refusal regression, introduces another recurring refusal or real
workflow regression above the review threshold, changes ownership/error/report
identity, increases code or stack materially without compensating benefit, or
fails any required gate. Rejection preserves the measured candidate witness
and semantic tests, restores only the production codec, and records the exact
candidate-to-restored transition. No production performance claim follows
from this experiment.

## ADR and goal constraints

The static design follows the accepted records already bound by the packet:

- ADR 0001, 0002, and 0024 keep the helper private to the existing shared
  OOXML owner with no public type or dependency edge.
- ADR 0003 is unchanged: immutable snapshots, edits, commits, patches,
  conflicts, and exact no-op publication remain downstream contracts.
- ADR 0005 and 0031 keep the scan and namespace view within existing source,
  output, and execution policy. No retained cache, ambient I/O, worker, or
  global pool is introduced.
- ADR 0006 keeps byte-oriented lexical dispatch, preservation, validation,
  typed refusal order, and malformed-input defenses authoritative.
- ADR 0008 requires frozen source identities, differential preservation
  evidence, adversarial controls, assembly/resource review, and complete
  repository gates before retention.
- ADR 0010 and 0011 keep facade, archive, physical package, and relationship
  ownership unchanged.

No accepted ADR is amended or proposed by this experiment. The candidate is a
private representation change whose disposition depends on the fresh 0702
evidence.
