# ADR 0003: Snapshots, edits, patches, and concurrency

- Status: Accepted
- Date: 2026-07-31

## State model

Opened documents are immutable, cheap-to-share `Send + Sync` snapshots. Major
objects are small lifetime-free handles backed by hidden shared state; fine
traversal uses borrowed views. Borrowed zero-copy inputs produce scoped
`DocumentRef<'a>`-style types, and conversion to owned storage is explicit.

Mutation consumes or borrows a snapshot into an `Edit`. Attached document trees
are never publicly mutable. Transaction-scoped proxies expose short semantic
verbs such as `set`, `add`, `move_before`, `clear`, and `remove`. Detached
builders may be ordinary mutable values when every field combination is valid.
Identity-changing verbs such as `rename` update their modeled dependency
closure in the same transaction. They never expose a catalog-string-only mode
through the ordinary facade.

`commit()` validates the changed dependency closure atomically and returns a
named `Commit<T>` containing the new snapshot, a reversible patch, and
diagnostics. It never mutates the source snapshot. A full validation pass is
explicit. Untouched malformed or unknown content is preserved and may remain a
diagnostic, but a commit cannot worsen it.

IWA reference-line edits apply the same boundary at the archive adapter:
malformed existing line payloads are rejected before patching, recognized
fields are checked for canonical framing, nested custom-value unknown fields
survive scalar replacement, and repeated graph nodes are bounded before
generated protobuf materialization. Public format editors publish only after
their staged CRUD operation and typed readback succeed.

## Identity and selection

Handles carry snapshot lineage and stable internal identity. Public lookup is
selector-first:

```rust,ignore
let sheet = book.sheet("Summary")?;
let first = book.sheet(0)?;
```

Both calls return `Result<Option<_>>`; numeric positions are checked and
zero-based. Names, semantic roles, A1/R1C1 references, and relative operations
are the main entry points. Raw physical IDs are advanced diagnostics only.

Structural insertion uses the same selectors as stable source-snapshot
anchors. `add_before` and `add_after` resolve a developer name or checked
zero-based position to semantic identity; they do not accept relationship IDs,
native sheet IDs, or part names. Base-object moves are resolved first,
before/after additions then stay attached to their anchor identity, repeated
additions retain explicit call or join order, and unanchored `add` remains the
concise tail operation. A transaction-local handle may report its current
projected position, but a later structural intent can shift it; the committed
patch is authoritative for final positions.

Serializable patches do not inject private IDs into Office files. They anchor
objects using available native identity, parent and semantic selectors, context
fingerprints, expected-state hashes, and patch-local IDs for inserted objects.
Ambiguous resolution is a conflict.

## Patches and conflicts

Every edit records semantic operations with read/write sets and yields a
versioned, format-independent patch. `Patch<Reversible>` has an inverse; a
consuming `seal()` creates `Patch<ForwardOnly>` for redaction and other
irreversible work. The canonical wire representation is deterministic JSON,
with optional compact encoding and content-addressed bounded blob bundles.

Independent `SubEdit`s may run concurrently. Join automatically merges only
provably disjoint effects. Overlap returns a structured `ConflictSet`; there is
no last-writer-wins behavior. Three-way merge produces a non-mutating plan and
commits only after every conflict is resolved.

Cross-document copying and moving use dependency-closure transfer primitives.
Move coordinates two snapshots and returns both results and patches. Resources
are reused only when equivalence is proven; name/style/theme collisions require
an explicit resolver. Undo/redo is an explicit budgeted `History<T>`.

## Deletion

`clear` removes primary payload while preserving the object. `remove` deletes
the selected object and required references. `gc` is a separate planned
reachability operation. Structural edits update unambiguous dependencies and
return a disposition requirement when content or incoming references would be
lost. Until a disposition is explicitly supplied, the safe facade returns a
typed refusal instead of guessing whether to cascade, detach, retarget, or
retain an orphan. Sanitization certification is available only when all
prohibited content is proven absent.
