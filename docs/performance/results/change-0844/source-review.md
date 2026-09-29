# Independent applied-source review

The reviewer inspected the complete applied `members.rs`, not just patch
context, and found no source-contract blocker. The first-missing borrowed
lookups follow all original conflict checks. Prefix closure justifies staging
only the missing suffix: every previously accepted name contributed all of
its proper ancestors. The suffix receives one exact Vec reservation, followed
by each owned String reservation. The original names, by_folded, and final
`descendant_by_folded_ancestor.try_reserve(ancestors.len())` reservations still
precede every logical mutation. No per-key map reservation remains.

Tests cover zero initial capacity with deep/partly-shared/distinct names,
case-folded shared ancestors retaining the first descendant, and all four
conflict forms leaving every index unchanged before a valid insertion.
Allocator failure injection is not claimed; the reservation and mutation
ordering is directly reviewed. Actual quality results are separately retained.

Pure root/shared/deep-shared/deep-unique fixtures bound the local tradeoff.
They are not a representative semantic Office corpus. A mixed catalog would
be useful additional evidence, but absence is not a source-correctness blocker;
adoption claims must remain scoped to the measured substrate operations.
