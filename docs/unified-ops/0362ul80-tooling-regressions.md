# 0362ul80 — tooling-row regression repairs

This source-only leaf is based on integration commit
`0dc6c76634555aa610d31f43ca96b739f93b8720`.  It repairs three deterministic
Python tooling regressions found by the first current-source tooling-row run.
No Cargo, rustdoc, native qualification, allocator, timing, or performance
run was performed for this leaf.

## Preserved failed current-source receipt

The actual 39-module `tooling-unit` row was run against `0dc6c766...` with
the current tracked runner (`e88893d497f7e7ef70bde4314c3e577d47f20997f2ea9d82b4d4595edbbe8458`),
not the historical f5 runner (`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`).
It ran 1,745 tests in 657.341 seconds and failed with two failures and one
error.  The receipt remains external and is not relabelled as a pass:

- `tools.test_native_qualification_runner.NativeQualificationRunnerTests.test_expired_compile_lease_reaps_child_and_writes_failed_receipt`
  raised `FileNotFoundError` because `output/events-build.jsonl` was absent;
- `tools.test_no_3rdparty_references.NoUpstreamCheckoutReferences.test_tracked_code_does_not_reference_the_upstream_checkout`
  reported the three source-policy predicates at
  `prepare_signed_native_inputs_v2.py:262,347,708`, the raw-coverage predicate
  at `raw_api_coverage.py:98-99`, and the negative fixture at
  `test_signed_native_policy_v2.py:279`;
- `tools.test_ink_safe_runner.SafeInkRunnerTests.test_fast_exit_tmp_cap_is_failed`
  observed return code `-15` instead of `0`.

The wrapper is
`(deleted-coord-artifact)/tooling-row-0dc6c76634-20261025T-wrapper.log`
(4,141 bytes, SHA-256
`f30aeca725ac3b8e44c3fb567e55564f1b7001d835ee9b472a52cd3ac28151db`).  Its
child log is
`(deleted-coord-artifact)/48-tooling-unit-attempt-1792900092755500553-677681.log`
(639,430 bytes, SHA-256
`222fb2ce0dacd5cf44b7852c4ab7b3c6cf9dd04ac97642b05aaa4d4e871e3213`).
The failed run ended with target size zero, an empty Cargo-target trace,
empty owned-process lists, unchanged source/lock identities, and no live
guard errors.  It was preserved without retry.

## Repairs

The no-upstream guard now has exact, line-scoped exceptions only for policy
predicates and the reviewed negative fixture.  A file-wide allow-list is not
used.  Stale/duplicated exception lines and path-bearing decoys are tested;
actual upstream fixture paths remain rejected by the policy tests.

The lease-expiry test now uses a five-minute permit plus an injected clock and
a process-start observer.  The clock advances only after the real fake Cargo
child is bound, while the permit file/hash remains unchanged.  The child is
held without a duration, the monitor validates the expired lease, reaps the
child, and the event journal and failed receipt are asserted.

The Ink stop path rechecks the child after the asynchronous unit kill before
sending a process-group signal, preserving a naturally completed return code.
The fast output/TMP cap tests use a marker handshake so the cap observer sees
a live child and the fake unit stop waits for its clean exit; they still assert
failed cap status, cleanup, and return code zero.

## Focused verification

```text
PYTHONPATH=tools python3 -m unittest -v \
  tools.test_ink_safe_runner \
  tools.test_native_qualification_runner \
  tools.test_no_3rdparty_references
Ran 60 tests in 20.707s — OK

python3 -m py_compile \
  tools/test_no_3rdparty_references.py \
  tools/test_native_qualification_runner.py \
  tools/ink_safe_runner.py \
  tools/test_ink_safe_runner.py
git diff --check — OK
```
