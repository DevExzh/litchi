# 0549 candidate source review

This is a read-only review of the isolated 0549 CFB candidate. The candidate
source snapshot is `candidate-sources/file.rs`, SHA-256
`711572fd5779aff9efdfd3e452040daa378e23aad22bb6ecb83841dbbe903377`. It is
based on repository revision
`6d9fbb7401729aacf2afc5d6b6c681a9e7384056`; the live production
`crates/litchi-cfb/src/file.rs` is SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
prepared `candidate.patch` is SHA-256
`0602a50efe65c800ea36163d22af19d4f44dac9a77f4d87861230956a23fb40d`, and
the bound source note is SHA-256
`135628353a35fd53cbb72b91914530454bac279796bdda0ee152727489929660`.

No production source was changed by this review. I ran no Rust build, test,
benchmark, profile, hardware-counter capture, allocation capture, or guard
run. The candidate remains isolated until the coordinator completes the
frozen baseline and candidate campaign.

## Static disposition

No source-level correctness or resource-boundary blocker was found. The
candidate is admitted to the isolated measurement campaign only. Runtime
retention still requires the exact candidate source and binary to pass the
owned-source differential and error contract, malformed-input guards,
allocation and workflow gates, profile and native gates, and final quality
checks in `plan.json`. This review does not admit production adoption. Any
change to the bound source, patch, or source-note bytes invalidates this
disposition and requires a new binding review.

The production diff has exactly one file and one hot call-site substitution:
it adds the private `CheckedBitSet::test_and_mark`, uses it in
`SectorChainScratch::collect_exact`, and adds private tests. The existing
`contains` and `insert` implementations and every other caller remain
unchanged. There is no public API, dependency, unsafe-code, limit, reader,
ownership, or ODF change.

## Combined operation and bounds

`test_and_mark` preserves the two fallible checks of the old
`contains`/`insert` sequence. It first rejects `bit >= bit_len` with the same
capacity diagnostic, derives the same word and mask, and then performs one
checked mutable word lookup with the same missing-backing-word diagnostic.
For an existing word it saves `old_word`, computes `old_word | mask`, stores
that value, and returns whether the old mask was set. The success path has no
reserve, resize, collection, or error-string allocation.

`collect_exact` still resets the scratch state, validates an empty
declaration, start marker, and declared count, reserves `sectors` with the
existing `"sector-chain entries"` resource label, and prepares and zero-fills
the visited map with `"sector-chain map"`. Inside the walk it performs the
same `u32` to `usize` conversion and allocation-table bound check before the
visited operation. The new helper therefore sees a slot that is in the
prepared bit-set range under the established invariants. It is followed by
the same sector append, table lookup, intermediate marker checks, and final
`ENDOFCHAIN` check.

The mutable table is only borrowed as `&[u32]` for this call. On a first visit,
the new store is exactly the old `insert` state transition. On a duplicate,
the bit is already set, so storing `old_word | mask` leaves the word's value
unchanged. `SectorChainScratch` is private and there is no alias or callback
between that store and the cycle error/reset, making the extra write
unobservable through the scratch state or its current consumers.

## Error precedence and state cleanup

The preflight error order is unchanged:

```text
Empty {table_name} chain must start with ENDOFCHAIN
Invalid start marker for {table_name} chain
{table_name} chain length exceeds its allocation table
Invalid sector index {sector} in {table_name}
Cycle detected in {table_name} chain at sector {sector}
{table_name} chain exceeds its declared length
{table_name} chain ends before its declared length
Invalid sector marker 0x{next:08X} in {table_name} chain
```

The current-sector conversion and table bound check still precede duplicate
detection. A duplicate still returns the cycle diagnostic before the append,
successor lookup, or any late-length diagnostic. If the checked helper's
out-of-range or missing-backing branches were reached, their text matches the
corresponding old `insert` failure; the normal `collect_exact` invariants
make those branches unreachable after the same checks. The direct helper
tests cover both branches without relying on that invariant.

The outer result handler still calls `reset()` for every collection error,
clearing the logical sector vector and `visited.bit_len` while retaining
capacities. A failed call can leave bits in retained words, as the old walk
could; the next non-empty call executes the unchanged full `fill(0)` before
using them. A successful call marks every visited sector just as the old
two-operation sequence did. No partial sector slice is passed to a claim
operation after an error.

The formatted corruption branches in `test_and_mark` can allocate their
diagnostic `String`, just as the old fallible helper did. The candidate makes
no claim of zero allocation for error construction; it removes only the
redundant checked word operation on the valid scratch walk.

## Allocation and ownership review

The candidate preserves reservation order and labels. `sectors` is reserved
before `prepare_visited`, and the exact reservation still covers every
in-loop `push`. `test_and_mark` cannot grow either buffer and introduces no
new fallible allocation boundary. Its direct missing-word and out-of-range
errors leave the vector contents, pointer, and capacity unchanged.

`validate_stream_allocations` still owns separate `mini_scratch` and
`regular_scratch` values. It consumes only the immutable `sectors()` slice
after a successful collection for mini-sector ownership checks or physical
sector claims. Those borrows end before the next mutable collection, and no
visited map or mutable scratch reference escapes. The combined helper is
private to `CheckedBitSet`; all directory traversal, owned chain collection,
mini-sector ownership, and other bit-set users retain the original
`contains`/`insert` behavior.

## Candidate test coverage

The focused `test_and_mark` tests cover false then true results at bit 0, the
last bit of the first `u64` word, and the first bit of the next word. They
also cover the exact `bit_len` boundary, the missing backing word, the exact
diagnostic strings, and unchanged pointer/capacity state after successful and
failing calls. The expected first word is the two-bit value
`1 | (1u64 << (BITSET_WORD_BITS - 1))`; it does not incorrectly require
`u64::MAX` when only bits 0 and 63 were set.

The retained scratch differential and reset tests compare valid, empty,
cycle, invalid-start, invalid-index, early-end, late-end, invalid-marker,
short-table, and empty-nonterminal cases against the independent owned
`collect_sector_chain_exact` helper, including exact error text and reuse
after errors. The candidate also retains the exhaustive oracle: table lengths
zero through four, every valid slot plus an out-of-range slot and the
`ENDOFCHAIN`/reserved-marker values, representative starts, and counts
through one above the table length produce 376,264 table/start/count
combinations. Each result is compared for exact sector order and diagnostic
text; successful non-empty walks check visited membership and errors check
scratch reset state. It deliberately does not assert zero visited bits on a
successful collection because this candidate preserves the visited marks.

These tests are meaningful source-level evidence, but they cannot cover all
table lengths, allocator-failure injection, a 32-bit target, or the full
physical claim graph. The coordinator must therefore retain the complete
candidate correctness, malformed-input, allocation, workflow, profile,
native, and quality gates. Assembly attribution should match the exact
`SectorChainScratch::collect_exact` owner and keep the unchanged
`CheckedBitSet::insert` callers separate from the inlined combined operation.

The candidate is consequently suitable for measurement only. OLE2 and OOXML
remain the active optimization priority, with ODF deferred until that goal
completes.
