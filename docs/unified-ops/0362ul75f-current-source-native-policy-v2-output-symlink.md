# 0362ul75f — reject output-artifact symlink escapes

Python/source-only follow-up to 0362ul75e.  No Cargo, native probe, rustdoc,
timing, allocator, or performance execution was performed.

Output-artifact validation now checks raw lexical path components before any
resolution.  A symlinked descendant such as
`output-root/target -> /outside` is rejected even when the resolved file has
the expected relative route name and matching bytes.  The validator also
requires lexical containment beneath the prepared root and rechecks the final
regular file identity at use time.

Added a regression covering a prepared-root `target` symlink escape.  The
existing absolute/traversal execution-receipt tests remain in the suite.

Focused verification:

```text
PYTHONPATH=tools python3 -m unittest tools.test_current_native_policy_v2
```

Result: 18 tests passed.  No native or Cargo work was performed.
