# 0362ul135e — f2 PPT forward-replay fidelity and tagged-route hardening

Date: 2026-10-25. This is a source/Python-only correction to 0362ul135d. No
Cargo, rustdoc, native probe, timing, allocator, or performance workload was
run.

The reviewed f2 PPT adapter now serializes the forward replay result, compares
its bytes with the committed target bytes before writing a receipt, emits
`replay_target_sha256` and `replay_matches_target`, and refuses a mismatch.
The adapter identity is now 3,561 bytes, SHA-256
`3b7e581bf01f66f70e3fe21d9a3ebe18065206cbd1e00ff07d9e36b9ecae1e06`. The
baseline policy, manifests, and materializer all bind this new exact identity;
the original f2 source and archive remain unchanged.

`validate_ppt_correctness_receipt.py` and the v2 route consumer now require
both forward-replay fields and equality. In `_v2_result_payload`, a record
tagged `ppt-scoped-correctness` is still checked as a complete correctness
receipt, while tagged suite/build/validation records must carry their exact
successful status and validation replay relation. A different route tag or a
minimal route-tag shortcut is rejected. This closes the prior branch that
accepted a route label without the semantic payload.

Focused verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_current_native_policy_v2 tools.test_no_3rdparty_references
```

Result: 33 tests passed. Coverage includes positive and malformed replay
receipts, missing replay fields, route-tagged correctness records, minimal
tagged probe/suite refusals, and the existing paired source/archive/ABBA
checks. No native or performance claim is made.
