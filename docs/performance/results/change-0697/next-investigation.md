# 0697 next investigation — borrowed inherited namespace views

Status: review-only hypothesis after retained change 0696. No production
source or public API change is implemented by this note. `performance_claim: none`.

The highest-impact semantically narrow follow-up is to make the ephemeral
`Inherited` value borrow the two namespace owners it observes from the parent
`Frame`, instead of cloning both `Arc` owners for the duration of one
`start` call. The current profile gives this target a stronger direct signal
than a larger `Ctx`/`Frame` representation change: the candidate's
`drop_in_place<Inherited>` is 2.82% of sampled cycles, while
`drop_in_place<Ctx>` is 1.33% and `start` is 9.28% in the 0696
open-plus-capture candidate profile. The corresponding baseline values are
2.82%, 1.42% and 12.11%. These are profile shares, not an end-to-end speed
prediction.

The fresh 0697 isolated MCE profile strengthens the same narrow target: self
shares are `start` 26.52%, `drop_in_place<Inherited>` 7.16%, and
`drop_in_place<Ctx>` 4.86%, with zero lost samples. The run is bound to binary
SHA-256
`02312667a7b0931fa37501702315241c1b0c4e6954d87d49ab625098a6dddd75` and
profile SHA-256
`831d7428e147b6c3ae4308947f66e5abcf66976604f0230b00b6fe319302a073`.
The inclusive callgraph also attributes 22.02% to inlined
`NamespaceLayer` `Arc` clone regions: `Inherited::after` 8.71%, parent
`Ctx` cloning 8.37%, temporary inherited namespace observation 2.53%, and
temporary inherited emitted-boundary observation 2.41%. Those are diagnostic
sample shares, not exact atomic-instruction costs or an additive speedup
estimate; samples at a following branch can reflect skid from the preceding
region. The address-bounded review is in
[instruction-review.md](instruction-review.md).

## Current ownership and the redundant work

The relevant private types are in
`crates/litchi-ooxml-common/src/mce/codec.rs`:

* `Namespaces` owns `head: Option<Arc<NamespaceLayer>>` and a binding count at
  lines 311–315. Each `NamespaceLayer` owns its parent layer through another
  `Option<Arc<NamespaceLayer>>` at lines 317–320.
* `Ctx` owns a `Namespaces`, an `Option<Arc<DirectiveLayer>>`, and the opaque
  flag at lines 470–475.
* `Frame` owns a complete `Ctx`, the mode and active flag, and a separate
  `emitted_ns` owner at lines 602–610. The separate emitted boundary is needed
  when dropped compatibility wrappers or branches lie between an element and
  its nearest emitted ancestor.
* `Inherited` is an ephemeral pair of owning options at lines 439–444. It is
  used only by `write_start` and `after`.

At `start` lines 763–767, the parent `Ctx` is cloned into the child working
`Ctx`, then `Inherited` clones the same namespace head and the parent frame's
emitted boundary again. `Inherited::after` at lines 451–458 performs the
remaining owner clone that must be retained in the child `Frame`: an emitted
element gets its current namespace head, while an un-emitted element carries
forward the nearest emitted boundary. The 0696 guard at lines 781–783 removes
only the empty `with_local` clone-and-replace; it does not touch these
`Inherited` owners.

The borrowed candidate would change only the temporary shape conceptually to:

```rust
struct Inherited<'a> {
    ns: Option<&'a Arc<NamespaceLayer>>,
    emitted: Option<&'a Arc<NamespaceLayer>>,
}
```

The references would point to `parent.ctx.ns.head` and `parent.emitted_ns` in
the existing parent frame. `write_start` already treats both values as read
only: `Inherited::hoists` at lines 446–449 passes them to the pointer-aware
`hoists` helper, and `write_start` lines 1349–1381 passes them to
`Namespaces::for_each_hoisted`. `Inherited::after` would still clone exactly
the one `Arc` needed by the child frame. Thus this hypothesis removes the two
temporary `Inherited` increment/decrement pairs without changing the
namespace layer graph, the nearest-emitted invariant, `Ctx` ownership, or
output bytes.

No allocation reduction should be assumed: the candidate removes reference
count traffic, not `Vec` or `String` allocations. The source profile and the
0696 allocation matrix must be treated separately.

## Lifetime proof and the Alt seam

The borrow must be taken from the parent frame, not from `c.ns.head`. The
working `c` is cloned at line 763 and may be replaced by
`with_local` at lines 781–783. A reference into `c.ns.head` would prevent that
replacement and would describe the wrong pre-local scope. The parent frame is
stable until `close` mutates `st`, so its fields are the correct source for
both inherited views.

The current order has one borrow hazard. Lines 981–986 take `st.last_mut()`
and mutate the parent's `Mode::Alt` counters and selection flags. A borrowed
`Inherited` created before that point would keep an immutable borrow of the
same vector element alive and make the later mutable borrow invalid. The safe
seam is therefore:

1. Clone the parent `Ctx` and complete all namespace/directive validation in
   the existing order.
2. Compute `parent_active` as today.
3. Run the `st.last_mut()` Alt bookkeeping in a short scope and retain only
   its value result (`active` plus `Mode`). Do not construct `Inherited` yet.
4. After that mutable borrow ends, enter a short block that obtains
   `st.last()` and constructs the borrowed `Inherited` from the parent
   `Frame`.
5. In that same block, call `write_start` where the current path does, call
   `Inherited::after` to make the owned `Frame.emitted_ns`, and construct the
   `Frame` with owned `c` and `Mode` values.
6. Let the block end, or explicitly drop the temporary view, before calling
   `close(st, frame, empty, out)`. `close` may reserve and push into `st` at
   lines 1134–1148, so no reference into the vector may survive that call.

The opaque early return at lines 786–806 needs the same inner block, but it
does not pass through the Alt mutable borrow. The direct Alt-child return paths
at lines 991–1001 and 1051–1061 must first build an owned `Frame` inside the
block and only then call `close`. The ordinary emitted path at lines
1101–1131 follows the same pattern. This is a structural lifetime
requirement, not an assumption that non-lexical lifetimes will infer the
desired disjoint-field borrow through every return path.

The proposed borrowing does not move or mutate the parent `Ctx`,
`parent.emitted_ns`, or any `NamespaceLayer`. The Alt mutation changes only
the parent's mode state; delaying the namespace view until after that scope
preserves the current validation and selection order. On `empty` elements,
`after` must still be evaluated before `close` so the ownership and error
ordering remain the same as the current code, even though `close` consumes the
frame without pushing it.

The most important static checks for an implementation are:

* no `Inherited<'_>` reference, helper return, closure capture, or local alias
  reaches `close`;
* `st.last_mut()` is completed before the borrowed view is constructed;
* the borrowed `ns` is the pre-local parent head, while `after` sees the
  post-local child `c`;
* `Frame.emitted_ns` remains owned and continues to distinguish emitted from
  dropped frames;
* the root path still accepts `None` for both views without manufacturing a
  temporary owner; and
* no direct or indirect change is made to `Namespaces::with_local`'s duplicate,
  limit, or parent-link behavior.

The intended ownership shape is a block boundary like this (illustrative
pseudocode; it is not a production patch):

```rust
// All parent Alt mutation has already returned or ended above this point.
let frame = {
    let parent = st.last();
    let inherited = Inherited {
        ns: parent.and_then(|f| f.ctx.ns.head.as_ref()),
        emitted: parent.and_then(|f| f.emitted_ns.as_ref()),
    };
    if emitted {
        write_start(/* ... */, &inherited)?;
    }
    Frame {
        emitted_ns: inherited.after(&c, emitted),
        ctx: c,
        mode,
        active,
    }
}; // all references into st end here
close(st, frame, empty, out)
```

The opaque and direct Alt-child paths use the same block but their existing
`write_start` and `active` decisions. A real implementation must preserve the
current early-return/error branches around this shape; the block is the
ownership proof boundary, not permission to reorder validation.

## Why a `Ctx`/`Frame` redesign is a separate, larger candidate

The `Ctx` clone at line 763 cannot be removed by the same local borrow. The
child may install a new namespace layer, replace its directive layer through
`c.directives.take()` at lines 953–960, and set its opaque state independently
of its parent. The parent frame must remain available for later siblings and
for the matching end event. Avoiding that clone would require a delta or
copy-on-write frame representation and a new proof for all error and depth
paths.

Likewise, `Frame.emitted_ns` is not simply a duplicate of the immediate
parent's `Ctx`: for skipped, unwrapped, Alt, Choice and Fallback frames, the
nearest emitted ancestor can be several stack entries away. Removing its
owner would require encoding that invariant in a different frame state and
rechecking `hoists`, `after`, branch selection and every end-event pop. The
current borrowed `Inherited` candidate leaves this representation unchanged
and therefore has a materially smaller semantic surface.

## Hazards and required evidence

The principal hazards are a wrong pre-local scope, an accidentally live
borrow across `close`, or an altered Alt branch selection/hoisting boundary.
An implementation must retain exact output bytes, ownership (`Cow` borrowed
versus owned), complete `Report`, and exact typed refusal identity. The
following cases are required in addition to ordinary formatting and focused
tests:

* root and nested elements with no declarations, inherited declarations,
  `xmlns=""`, prefix rebinding, and shadowed hoisting;
* `AlternateContent` with selected and rejected `Choice`, `Fallback`, nested
  branches, invalid branch order, and non-ignorable children;
* `ProcessContent`, `PreserveElements`, `PreserveAttributes`, opaque
  descendants, and dropped wrapper scopes where `emitted_ns` differs from the
  immediate context;
* empty elements, multiple-root and late-error paths, exact namespace/depth/
  directive/choice limits, and QName-before-directive refusal order;
* the complete existing shared oracle (all real, mutant and synthetic cases
  under every capability profile), including output SHA/length, `Cow`
  ownership, `Report` and `Debug` error identity;
* the refusal/control matrix and the focused MCE suite, followed by the same
  native end-to-end and declaration-heavy controls used for 0696; and
* frozen assembly/instruction attribution showing the `Inherited` drop is
  absent or reduced while `Ctx`/`Frame` ownership and stack behavior remain
  bounded. Allocation counts and requested bytes should be compared even
  though no reduction is expected.

The candidate should be rejected if any source-bound oracle, refusal identity,
report counter, ownership result, or limit/error order changes. A measured
regression in declaration-heavy inputs must remain visible even if the common
real-deck path improves; the 0696 control already established that this tradeoff
cannot be silently classified as noise.

## ADR and ownership matrix

| Constraint | Review consequence |
| --- | --- |
| ADR 0001, priorities and API layers | This is private codec plumbing; no public selector or API layer changes. Correctness remains ahead of the refcount saving. |
| ADR 0002 and ADR 0024, crate ownership/topology | The namespace graph remains owned by `litchi-ooxml-common`; no archive dependency, facade shortcut or new shared global is introduced. |
| ADR 0003, snapshots and concurrency | `Frame` and `Ctx` remain private immutable parent-linked state during one parse. No public snapshot, `Send`/`Sync`, mutation, or publication contract changes. |
| ADR 0005, I/O, memory and performance | The borrow is stack-scoped and bounded. It removes temporary `Arc` traffic without weakening limits or claiming allocation/RSS savings; native evidence must include tails and uncertainty. |
| ADR 0006, preservation and validation | Namespace resolution, output, reports, refusal order, validation and unknown markup must remain exact. No repair, normalization, or limit relaxation is allowed. |
| ADR 0008, migration and verification | The source diff, oracle, focused tests, integration gates and evidence bindings must prove the ownership seam before retention. |
| ADR 0031, execution budgets | No execution context or budget accounting is changed; any fragment/namespace work remains under existing limits. |

The recommended next experiment is therefore the borrowed `Inherited<'a>`
view, with no simultaneous `Ctx`/`Frame` redesign. Its expected value is a
direct test of the largest measured destructor target left after 0696, while
its semantic proof can be kept local to `start`, `Inherited`, `write_start`
and `close`.
