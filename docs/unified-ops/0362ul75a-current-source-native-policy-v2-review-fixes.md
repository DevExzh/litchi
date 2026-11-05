# 0362ul75a — current-source policy v2 trust-boundary fixes

This is a source/Python-only follow-up to 0362ul75.  No Cargo build, native
probe, rustdoc, timing, allocator, or performance claim was made.

The review fixes close the following receipt and execution gaps:

- source observation now parses the complete `HEAD` tree, rejects sparse
  `skip-worktree` entries, verifies every tracked regular blob against its Git
  blob, and rejects dirty, untracked, ignored, symlinked, or overlay paths;
- preparation rejects ancestor/descendant endpoint roots and overlaps among
  source, output, and optional source-archive paths;
- the PPT v2 plan binds `x86_64-unknown-linux-gnu`, `release`, and
  `--all-features` consistently in build, test, and clippy commands, and
  guards source/target/state/archive disjointness;
- source archives have fixed size/hash pins and an explicit tar-to-Git-tree
  regular-member proof path, following the reviewed signed-native archive
  shape;
- the ignored root `Cargo.lock` is admitted only when it is the exact lock
  pin declared by the endpoint policy; every other ignored or untracked path
  remains rejected;
- preparation tooling and manifest path/hash identities are reloaded and
  checked at validation time, including policy/preparer/validator bytes;
- qualification receipts must bind the prepared source/lock/archive
  identities and exact route role/program/output/status metadata, including
  actual output-file identities under the prepared disjoint output root and
  per-route source bindings.  Refusal routes cannot masquerade as successful
  build/open/text routes.

The candidate remains incomplete: abbreviated font references are still
provenance hints only, and no execution is admitted until exact candidate
source/tree/lock/probe/fixture/archive pins are reviewed and committed.

The updated pending-policy digest is
`d380c25389b289b0edc774e3932d04055b7f717bcb80a666fab4be90d97241e4`.

Focused verification:

```text
PYTHONPATH=tools python3 -m unittest tools.test_current_native_policy_v2
```

Result: 15 tests passed, including sparse omission, ignored overlay, exact
pinned root-lock admission, tracked symlink, root-overlap,
source/output/archive overlap, manifest/tooling mutation, route/status
tampering, unrelated-output rejection, and pending-policy fabrication
refusals.
