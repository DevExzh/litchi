# 0362ul102 — signed-batch owner decision

Date: 2026-10-25.  The owner confirmed the recommended choice from the
coordinator contract checkpoint:

> Keep immediate refusal; correct the outdated equivalence test.

This closes the signed-batch contract question without changing production
admission.  A signed `File::batch` remains a protected write target.  Its
preflight calls the source-freshness fence and then returns
`DocumentProtected` before requesting or consuming an item, so its
`Error::item()` is `None`.  A refusal raised while a sequential `set` or an
ordinary unsigned batch reads an item remains item-scoped.  The two cases are
not equivalent and must not be normalized to the same item index.

## Consumer migration

`crates/litchi-ops/tests/pptx_batch_equivalence.rs` now preserves the error's
actual item scope instead of converting `None` to `Some(0)`.  The recursive
corpus walker includes producer-signed XMLDSIG decks, so the test branches on
the actual owner: unsigned corpus and refused-body entries still require
batch/closure/sequential equivalence, while signed entries assert the distinct
admission contract explicitly:

* signed batch: `DocumentProtected`, `item = None`, source bytes unchanged;
* one sequential signed setter: `DocumentProtected`, `item = Some(0)`;
* a one-shot iterator supplied to the signed batch is not advanced at all.

The corpus branch does not force every individual signed setter to be a
protected-document error: an invalid target still resolves through the
ordinary route first.  The focused `/slide[1]/notes/paragraph[1]` case proves
that a signed batch remains `DocumentProtected`/`None` without path resolution,
while both sequential `set` and closure `edit` retain their `PathNotFound` /
`Some(0)` result.

For multi-item signed closures, the individual route may resolve a later
invalid target after an earlier valid change has been staged.  The focused
shape-then-invalid-notes case therefore expects sequential `set` to stop at
the first signed change (`DocumentProtected`/`Some(0)`) while one closure
reaches the later target (`PathNotFound`/`Some(1)`).  The random signed corpus
branch checks only the batch admission and source-preservation contract; it
does not pretend these distinct individual route schedules are equivalent.

The iWork equivalence consumer already preserved `None` through its
`error.item().or(fallback)` helper and required no change.  No production
crate, facade admission path, or frozen qualification probe is changed.

The signed exact-no-op commit/replay, cancellation, signature-preserving
publication, refusal, unsign, and source-freshness contracts remain those of
[0362w](0362w-signed-noop-replay.md) and the shared signature tests.  This
decision does not add a batch route to the signed-native timing matrix.

## Qualification binding

The v2 signed-native manifest binds the decision as
`signed_batch.status = owner_confirmed` and
`signed_batch.selection = immediate_refusal`.  This records the contract
selection; it does not freeze the still-pending candidate source or relabel
0362ul24 receipts.  A future qualification must still prove, before any
timed route:

1. signed batch refusal occurs before stream consumption and has no item;
2. source bytes remain unchanged and no output is published on refusal;
3. unsigned/per-item equivalence retains the actual item index;
4. signed exact no-op commit/replay remains byte- and signature-preserving;
5. cancellation, changed signed edits, unsign-and-write, and source-freshness
   keep their existing typed outcomes and publication fences.

The signed-native route matrix remains limited to its existing leaf/facade
no-op, replay, refusal, unsign-write, unsigned-negative, freshness, and
budget routes.  No batch result is timed or used as native performance
evidence.

## Validation boundary

Source-only validation for this change is:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_signed_native_policy_v2
```

After integration, the owner must run the focused Rust checks for
`pptx_batch_equivalence` and the existing signature tests, including
`signed_batch_refuses_before_item_admission`,
`signed_pptx_shape_setter_preserves_exact_noops_and_refuses_changes`,
`invalid_intermediate_signed_edits_return_their_typed_error_without_publication`,
and `unsigned_revision_keeps_source_freshness_for_reads_edits_and_save`.
No Cargo, compiler, native process, profiler, or timing workload is part of
this source-only change.
