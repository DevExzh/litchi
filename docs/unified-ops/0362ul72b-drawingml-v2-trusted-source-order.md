# 0362ul72b — DrawingML v2 trusted-source ordering repair

Date: 2026-10-25. This source-only repair closes two fail-closed gaps found
in the independent v2 review. It does not run Cargo, rustdoc, a native probe,
or a timing workload.

The v2 preparer now verifies the fixed `472b...` ref, `fbc55c...` Git blob,
`280d61...` SHA-256, and 3,088,216-byte payload before copying any candidate
probe or route file into the baseline. A missing or corrupt trusted source
therefore leaves the baseline untouched. The preflight payload is passed into
materialization, so the source identity is not silently reread after a failed
preflight.

At receipt use, baseline identity validation now requires the materialized
`tools/perf-baseline/src/lib.rs` to match all three fixed identities: SHA-256,
byte count, and Git blob SHA-1. Receipt validation also rechecks the trusted
ref/blob through the supplied trusted Git root; a tampered historical file or
crafted matching receipt is rejected.

Negative source-only tests cover both order and tamper cases. The combined v1
frozen-source and v2 policy suite passes 26 tests:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_drawingml_raw_qualification \
  tools.test_drawingml_raw_policy_v2
```

No native, fidelity, allocator, RSS, or performance result is claimed.
