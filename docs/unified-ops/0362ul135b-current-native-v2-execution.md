# 0362ul135b — current-native v2 execution and receipt consumer

Date: 2026-10-25. This source/Python-only follow-up preserves 135 and 135a.
It does not run Cargo, rustdoc, a native process, timing, allocator, or
performance workload.

The v2 plans now use the exact fixed route IDs. PPT build/correctness steps
are named `ppt-scoped-build`, `ppt-scoped-correctness-suite`,
`ppt-scoped-correctness`, and `ppt-scoped-correctness-validate`. The legacy
plan retains all eight route IDs and invokes the refusal-contract consumer
without `--check-only`; the old one-line `{routes:8}` probe path is gone.

`current_native_policy_v2.execute_v2` is the reviewed v2 consumer. It first
revalidates the preparation receipt, runs fixed build steps, invokes route
steps in `before, after, after, before` order, writes one bounded stream per
route/ABBA leg under the prepared output root, writes the fixed
`execution-receipt.json`, writes a qualification receipt, and consumes that
receipt through `validate_current_native_v2.validate_qualification`. The
validator now binds the v2 runner identity, exact route receipts, output
artifacts, source bindings, and all four ABBA legs. The generic historical
runner receipt remains accepted for its existing lanes.

The executor is injectable for source-only verification and defaults to a
strict argv subprocess executor for a later quiet-host run. No caller can
change route IDs, source bindings, output-root paths, or ABBA order.

Focused verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
```

Result: 23 tests passed. The suite includes real-shaped injected-executor
receipts for all four PPT routes and all eight legacy routes, four ABBA legs
per route, route-ID/stream checks, and the existing source, lock, archive,
manifest, output, and protocol-boundary refusals. No native acceptance or
performance claim is made.
