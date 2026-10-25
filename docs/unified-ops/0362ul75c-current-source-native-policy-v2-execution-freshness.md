# 0362ul75c — current-native execution freshness and route-status binding

Python/source-only follow-up to 0362ul75b.  No Cargo, native probe, rustdoc,
timing, allocator, or performance execution was performed.

The v2 boundary now requires each qualification result status to equal its
fixed route status; successful build/open/text routes cannot be reported as
refused, and refusal routes cannot be reported as successful.  Qualification
outputs must be actual files at the exact policy-relative path beneath the
prepared output root, not an unrelated file with a copied relative label.

Preparation can create an output root only when the requested path is absent,
regularly parented, and disjoint from both source endpoints, archives, and the
preparation receipt.  It records the empty-root creation proof.  Qualification
then requires the actual `execution-receipt.json` under that root, bound to the
preparation receipt, policy, runner bytes, route metadata, source/lock/archive
identities, and the exact route evidence.  A stale pre-populated root or
fabricated all-refused route set therefore fails closed.

Focused verification:

```text
PYTHONPATH=tools python3 -m unittest tools.test_current_native_policy_v2
```

Result: 17 tests passed, including final-boundary valid output binding,
all-refused status tampering, stale output-root rejection, and runner receipt
identity checks.
