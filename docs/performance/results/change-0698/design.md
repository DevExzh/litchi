# 0698 — borrow inherited MCE namespace views

`performance_claim: none`

This packet evaluates one private codec candidate: make the ephemeral
`Inherited` value borrow the namespace owners already held by the parent
`Frame`, instead of cloning two `Arc<NamespaceLayer>` values for the duration
of one `start` call. The candidate must leave the parent `Ctx`, the child
`Frame`'s emitted boundary, and the namespace-layer graph owned exactly as
they are today. This note is a design and review record; it does not itself
claim a speedup or authorize a broader representation change.

## Baseline and narrow target

The baseline is commit `8f16c57b27fd7878ece718e22db80839583f0098`.
`Inherited` currently owns:

```rust
struct Inherited {
    ns: Option<Arc<NamespaceLayer>>,
    emitted: Option<Arc<NamespaceLayer>>,
}
```

At the beginning of `start`, after the parent `Ctx` is cloned, the two values
are cloned from the parent context and parent frame. `write_start` reads them
only to decide whether declarations must be hoisted and which bindings to
emit. `Inherited::after` then clones exactly the owner needed by the child
`Frame`: the current child namespace head when the element is emitted, or the
nearest already-emitted boundary when the element is dropped or unwrapped.

The candidate may change only the temporary view to an explicitly borrowed
shape such as:

```rust
struct Inherited<'a> {
    ns: Option<&'a Arc<NamespaceLayer>>,
    emitted: Option<&'a Arc<NamespaceLayer>>,
}
```

The candidate must not remove the parent `Ctx` clone or the owned result of
`Inherited::after`. Those values survive beyond the call's temporary view and
are required for local declarations, directive inheritance, opaque state,
sibling processing, dropped wrappers, and end-event behavior.

## Lifetime and ordering proof

The borrow source is the parent frame, not the cloned child context. The child
context can be replaced by `with_local` after the raw namespace declarations
are checked. Borrowing from `c.ns.head` would therefore describe the wrong
pre-local scope and would also interfere with replacing that field.

The implementation must preserve this order:

1. Decode attributes and validate namespace syntax in the existing order.
2. Clone the parent `Ctx`, collect local declarations, install them into the
   child context, and perform all QName/directive/limit validation exactly as
   today.
3. Finish every mutable borrow used for `AlternateContent` bookkeeping,
   including the `st.last_mut()` choice/fallback counters and selection flags.
4. Construct the borrowed `Inherited` view from `st.last()` (or `None` at the
   root) only after that mutable borrow ends.
5. Run the existing `write_start` call at the same semantic point, call
   `Inherited::after` to obtain an owned `Frame.emitted_ns`, and construct the
   complete owned frame.
6. End the borrow block before calling `close`; no reference, alias, closure
   capture, or helper return derived from the view may reach `close`, which may
   reserve or push into `st`.

The opaque path, direct AlternateContent child paths, and ordinary path all
need this owned-frame-before-`close` boundary. A source rewrite that relies on
non-lexical lifetime inference without making this boundary evident is not a
sufficient proof. Empty elements must still evaluate `after` before `close`,
including when `close` consumes the frame without pushing it.

## Semantic invariants

The following must remain identical for every source and capability profile:

- the inherited namespace head is the parent's pre-local head;
- local declarations still create a new parent-linked layer, reject duplicate
  declarations, enforce binding limits, and preserve the implicit `xml`
  binding accounting;
- `xmlns=""` remains a nonempty local declaration and takes the existing
  installation path;
- `hoists` and `for_each_hoisted` see the same pointer identities and nearest
  emitted boundary;
- `Frame.emitted_ns` remains an owned `Arc` and continues to represent the
  nearest emitted ancestor across skipped, unwrapped, AlternateContent,
  Choice, and Fallback frames;
- QName expansion, directive validation, opaque handling, branch selection,
  error precedence, limit checks, output bytes, complete `Report`, and
  `Cow` ownership remain unchanged; and
- no public API, cache, snapshot, layout, resource policy, concurrency,
  execution context, or crate dependency boundary changes.

The candidate must not change `Namespaces::with_local`, `hoists`,
`shadowed_before`, `for_each_hoisted`, `close`, or the streaming codec. Any
change beyond the private temporary view and the lifetime seam requires a
separate design review.

## Required review and measurements

Before retention, the packet must bind frozen baseline and candidate source,
build, lockfile, corpus, and binary identities. The shared MCE oracle must
cover real Office XML, deterministic mutations, and synthetic cases under all
capability/limit profiles, comparing exact output bytes and length, `Cow`
ownership, complete reports, and typed/debug error identity. The focused MCE
suite, refusal matrix, namespace/choice/depth/directive limits, and all
existing repository gates must pass.

The semantic matrix must include root and nested inherited scopes, empty
default resets, prefix rebinding and shadowed hoisting, opaque descendants,
ProcessContent/PreserveElements/PreserveAttributes, dropped wrappers where
the emitted boundary is not the immediate parent, selected and rejected
AlternateContent branches, empty elements, multiple roots, late errors,
QName-before-directive ordering, and exact limits.

Native timing, tails, allocation counts and requested bytes, profiles, code
size, stack reservation, and frozen before/after assembly must be reviewed as
separate evidence. The target is removal or reduction of the two temporary
`Inherited` clone/drop pairs. Any remaining parent `Ctx` and child-frame owner
operations must be explained. A representative end-to-end gain is required
for retention; no allocation or RSS saving should be inferred from borrowing
alone. Declaration-heavy controls must remain visible, including any
repeatable regression, and the +5% review trigger applies to every recorded
metric.

The final disposition is deliberately open until the frozen candidate
assembly, semantic oracle, and measured controls are reviewed together.
