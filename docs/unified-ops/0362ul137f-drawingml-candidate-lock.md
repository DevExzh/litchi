# 0362ul137f — bind DrawingML preparation to the current workspace lock

DrawingML's final source preparation refused frozen candidate C because its
v2 policy still expected the earlier workspace lock. The preparation module
and all three candidate-lock representations in `drawingml-raw-v2.json` now
use the actual approved C lock: 139,207 bytes, SHA-256
`02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a`.
The historical 139,606-byte baseline and both standalone role locks retain
their existing identities. The paired runner, validator and compile driver
read the shared policy and need no separate literal migration.

This change is confined to protocol P, based on
`53451cc8d22499c3532502b3c986339a7206afdd`. Production C remains
`f8cd398a770f7363f628be40edb3ce4219566c16`. Historical receipts retain their
original source and lock identities.

The new cross-consumer test checks the canonical candidate lock in the
preparation policy, nested manifest values and top-level fields, along with
the unchanged baseline. The external patch preparation also passes six
checks, including mutations of the nested/top-level candidate, baseline and
preparation values. Its apply-ready patch is preserved under
`.coord/native-final-bindings-20261025/drawingml-role-lock-v3/`, SHA-256
`381ccd254016b2c9158079dc6110a7171819ab331e2a49a6a59f872e3e84db9b`.

All 26 focused tests pass with:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest \
  tools.test_drawingml_raw_policy_v2 \
  tools.test_drawingml_raw_qualification_v2
```

Final candidate binding, archives, guarded native execution and measurements
remain required. This pin migration does not establish a native result.
