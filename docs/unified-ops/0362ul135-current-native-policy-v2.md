# 0362ul135 — current-native v2 role-bound preparation

Date: 2026-10-25. This is a source/Python-only protocol change. It did not
run Cargo, rustdoc, a native probe, a timing workload, or an allocator
workload. No production source pin is claimed: the final candidate source is
still pending the coordinator's immutable freeze.

## Policy boundary

`current_native_policy_v2.py` now models the endpoint identity as explicit
source manifests, three role-separated locks, toolchain/configuration, probe
inputs, fixtures, and an external Git-tree archive. The workspace, perf, and
native lock paths are fixed by the policy; a receipt or caller cannot select a
role path or supply a digest. The historical baseline retains the reviewed
workspace, perf, and native lock identities from the source policy. The
candidate role locks, source/tree/archive identities, probes, and fixture
payload remain incomplete until the final source is supplied.

Both `ppt-current-v2.json` and `legacy-current-v2.json` carry the same
derived trusted-input projection. Preparation and validation compare that
projection with the policy code, in addition to the exact ordered route
contract. A copied manifest, changed route, or candidate identity overlay is
therefore refused.

The PPT plan continues to bind the workspace lock, reviewed fixture, fixed
target/profile/features, and verified source archive. The new legacy plan
binds the perf lock and the current refusal-contract consumer, while retaining
all eight legacy routes (six successful and two refusal routes). Both plans
require the complete candidate source policy before returning executable
steps; pending policy remains diagnostic-only.

## Focused verification

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest tools.test_current_native_policy_v2
```

Result: 20 tests passed. In addition to the existing pending-policy and
receipt-hardening cases, the suite builds temporary Git-tree archives and
exercises a ready preparation through both consumer plans. It rejects wrong
source, role lock, archive, and route identities. No performance, fidelity, or
native acceptance claim is made here.
