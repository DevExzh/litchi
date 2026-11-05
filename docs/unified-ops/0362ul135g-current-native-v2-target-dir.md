# 0362ul135g — current-native v2 side-specific Cargo targets

This is a source/Python-only repair against frozen candidate C
`f8cd398a770f7363f628be40edb3ce4219566c16`. It must be applied in the later
protocol-only P commit and does not change the candidate source, source pins,
archives, locks, adapters, workloads, or route contract. No Cargo, rustdoc,
metadata, native, timing, allocator, or performance command is part of this
repair.

## Defect

`execute_v2` gives each endpoint an external target root and records binary
paths below that root. `make_ppt_steps_v2` and `make_legacy_steps_v2` passed
`--target` but omitted `--target-dir`, so Cargo wrote to the source checkout's
default `target/` while `_binary_identities` inspected the disjoint external
path. The source target cannot be reused: `_check_plan_paths` and
`_ensure_execution_paths` intentionally reject source/output overlap.

## Repair

Both builders normalize the caller target to an absolute path and put that exact absolute
path in Cargo's `--target-dir` option. The PPT build and all seven PPT
test/clippy commands use the same target directory; the legacy build uses its
own side-specific directory. Binary paths and target-dir values are therefore
the same reviewed endpoint identity for `before` and `after`. No environment
variable or caller digest is used as a substitute.

The existing ready-pipeline test now asserts the exact target-dir argv,
binary paths, and all seven PPT suite target-dir occurrences. It remains a
source-only injected test and does not invoke Cargo.

## Rebind requirement

The final P checkout must be clean after this repair, and its full commit plus
the repaired policy/preparer/validator hashes must be recorded. Only then may
the external f8cd binding recipe emit the final candidate pin and both
manifests. The pre-binding manifests and tooling identities from an earlier P
checkout must not be reused. Re-run the external packet's emitter with the
actual post-binding P commit, then prepare and validate both lanes from fresh
output roots.

Focused verification after applying this patch:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
```

Result: all 29 focused tests pass, and `git diff --check` passes.
This record makes no native readiness or performance claim.
