# 0362uoa — deferred iWork table-merge selector boundary

The acceptance tooling found that the Numbers and Pages table-merge edit
constructors intentionally defer rooted target lookup. Their constructors
still own and validate the typed selectors, while the first staged/read
operation performs the bounded source-backed resolution. The boundary checker
was still requiring the old eager constructor marker and therefore rejected
the correct lazy implementation.

The checker now distinguishes the two valid shapes. Eager families must retain
their selector-resolution marker in `edit_*`; lazy Numbers and Pages families
must expose a private resolver containing the typed target-resolution marker,
and every `regions`, `merge`, `unmerge` and `commit` entry point must call that
resolver. A marker in an unrelated helper cannot satisfy the check. The test
suite also removes each resolver call in turn and verifies that the topology
audit rejects the weakened source.

The production source is unchanged. This preserves O(1) edit entry and keeps
target resolution, source fencing and typed errors before any staged mutation
or publication.

Checks from the focused boundary packet:

```text
python3 -m unittest \
  tools.test_check_crate_boundaries.BoundaryPolicyTests.test_table_merge_transaction_boundaries_accept_checked_in_owners \
  tools.test_check_crate_boundaries.BoundaryPolicyTests.test_lazy_table_merge_transactions_guard_every_staging_entry
```

Both tests passed (2 tests, 6.622 seconds) on source `250769e1dc`.
