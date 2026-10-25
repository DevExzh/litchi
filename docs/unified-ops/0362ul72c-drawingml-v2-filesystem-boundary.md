# 0362ul72c — DrawingML v2 filesystem boundary hardening

Date: 2026-10-25. This source-only follow-up closes the final v2 preparation
boundary gaps. It does not run Cargo, rustdoc, a native probe, or a timing
workload, and it leaves v1 receipts unchanged.

The v2 preparer now rejects baseline/candidate paths that are equal or
ancestor/descendant, and rejects a preparation receipt path that overlaps
either source root. These checks run before source materialization, preventing
an endpoint or receipt from being nested inside another endpoint's tree.

Vendor validation now walks every materialized file below all three bound
vendor roots in addition to checking the Git tree. Any ignored or untracked
overlay, symlink, or sparse omission fails closed; legitimate tracked files
remain part of the exact vendor digest.

New negative tests cover nested endpoints, receipt-inside-source, and ignored
vendor overlays. The combined v1/v2 source-only suite passes 28 tests:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_drawingml_raw_qualification \
  tools.test_drawingml_raw_policy_v2
```

No performance, allocator, RSS, fidelity, or native acceptance claim is made.
