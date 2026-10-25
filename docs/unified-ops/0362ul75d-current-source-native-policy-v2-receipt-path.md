# 0362ul75d — bind the execution-receipt path to the trusted constant

Python/source-only follow-up to 0362ul75c.  No Cargo, native probe, rustdoc,
timing, allocator, or performance execution was performed.

The execution receipt is now always the fixed relative name
`execution-receipt.json`.  Preparation admission rejects any missing,
absolute, traversal, or otherwise substituted receipt-name field.  Execution
validation derives the path solely from the prepared output root and that
trusted constant; it never joins a caller-provided relative path.  Existing
symlink and output-root disjoint checks remain in force.

Added adversarial tests for both absolute and `../` receipt-path mutation.

Focused verification:

```text
PYTHONPATH=tools python3 -m unittest tools.test_current_native_policy_v2
```

Result: 18 tests passed.  No native or Cargo work was performed.
