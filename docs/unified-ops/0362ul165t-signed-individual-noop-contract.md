# 0362ul165t — Signed individual no-op test correction

Row 38 of the remaining-20 diagnostic built the Python binding and ran 301
tests. It passed 300 tests and failed only
`test_write_api.EditTests.test_protected_files_refuse`: the test expected
`signed.edit(lambda tx: None)` to raise `document_protected`, while the
binding returned `<Commit: changed=False, operations=0>`.

This is a stale assertion, not a production regression. The signed-edit and
signed-batch contracts are intentionally different:

* An individual `File::edit` closure may establish an exact no-op. The Rust
  transaction admission comment and final fence say that a signed closure is
  allowed to run, and a refusal is returned only when the completed edit is
  changed (`crates/litchi-ops/src/transaction.rs`, `admit` and
  `File::edit_with_options`). The facade then collapses that result back to
  the original signed file.
* The shared signature tests already require an empty signed edit to produce a
  no-op patch and to replay without changing signed bytes
  (`crates/litchi-ops/src/tests/signature.rs:1600-1617`). The signed policy
  also lists exact no-op commit/write/replay as successful routes
  (`0362ul74-signed-native-policy-v2.md:75-87`).
* A signed batch cannot prove a no-op before consuming its input. Its separate
  `preflight_batch` path therefore refuses immediately, before reading an
  item, as confirmed by the owner policy (`0362ul74-signed-native-policy-v2.md:59-60`)
  and the signed consumer contract (`0362ul136-signed-native-v2-consumer.md:37-40,77-79`).

The Python test now asserts the individual empty edit's `changed=False`, zero
operations, no-op patch, unchanged output bytes and retained signature. It
also replays that patch onto the signed source and checks the same invariants.
It then performs a genuine body mutation and retains the required
`ProtectedError` assertion. The existing signed-batch mutation refusal in
`test_signed_files_keep_their_signature_until_unsigned` is unchanged, so the
immediate batch refusal contract is not weakened or conflated with the
individual no-op policy. Locked-file refusal coverage is unchanged.

## Validation and handoff

No Cargo build, native qualification, archive or timing command was run during
the active diagnostic. The row-38 extension was preserved outside the active
target and exercised only by two focused Python tests:

* `/home/zhuhe/code/litchi-worktrees/.coord/remaining20-python-signed-noop-failure-20261026/liblitchi_py-row38.so`
  — 269,624,384 bytes,
  SHA-256 `b152e93729732deee82c53970a1888123e0004d312123a514ec62f5ffc1a641f`.
* Original failure log:
  `/home/zhuhe/code/litchi-worktrees/.coord/final-gate-current-remaining20-diagnostic-20261026-v3-state/38-python-bindings-attempt-current-remaining20-full53-38.log`
  — SHA-256
  `1ca770f5d3806db088a078783e3df856ec20b119fb891488410ef6fda3bada97`.

With the preserved extension, the direct no-build commands already passed:

```text
PYTHONDONTWRITEBYTECODE=1 LITCHI_TEST_DATA=/home/zhuhe/code/litchi-worktrees/uo2-signed-noop-test/test-data \
PYTHONPATH=/tmp/litchi-py-signed-noop.BvkCOI:/home/zhuhe/code/litchi-worktrees/uo2-signed-noop-test/crates/pyo3-litchi/tests \
python3 -B -m unittest -v test_write_api.EditTests.test_protected_files_refuse
# Ran 1 test in 0.006s — OK

PYTHONDONTWRITEBYTECODE=1 LITCHI_TEST_DATA=/home/zhuhe/code/litchi-worktrees/uo2-signed-noop-test/test-data \
PYTHONPATH=/tmp/litchi-py-signed-noop.BvkCOI:/home/zhuhe/code/litchi-worktrees/uo2-signed-noop-test/crates/pyo3-litchi/tests \
python3 -B -m unittest -v test_write_api.SaveTests.test_signed_files_keep_their_signature_until_unsigned
# Ran 1 test in 0.020s — OK
```

Root verifies that the only crate/tool change from the row-38 source `7c043ea168`
to leaf `85126247dc` is this Python test. The preserved extension therefore
exercises the unchanged production implementation. These two focused passes
are sufficient to validate the assertion repair; repeating their build or
execution is not required for this test-only merge.

The complete Python row, including its build and all 301 tests, remains
required in the fresh final-candidate gate. This record does not claim native
or full-gate acceptance, and the original failed diagnostic row is preserved.
