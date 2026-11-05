# 0362ul35 — current tooling coverage and marker-handshake repair

This source-only record separates the historical `0362uma` tooling result,
the interrupted frozen `9781` gate, and the current integration runner. It
does not claim a full gate, Cargo result, native result, or acceptance.

## Coverage boundary

The historical `0362uma` row 48 (`tooling-unit`) is a failed selected
diagnostic. It executed nine legacy modules and is not the frozen `9781` row.
The frozen `9781cf6d87e4d261d1dd9670e144e2db127d869b` gate stopped in rows
25–33, so its row 48 is `not_run`; it has no pass or failure result to reuse.

At current integration commit
`d07d30b1b646c42dcd2ffb763d2c06b8d29dd01e`, the runner had 25 tooling
modules. Commit `1079e43d5df347975b3ce80eedea5a0ef928cdff` registers six
additional current suites:

```text
docs/unified-ops/results/0362ul29/tooling/test_safe_chunked_tests.py
tools.test_ink_safe_runner
docs/unified-ops/results/0362ul24/test_prepare_signed_native_consumer.py
tools.test_opc_read_ahead_timing_guard
tools.test_xls_f3_qualification_abba
docs/unified-ops/results/0362ul24/test_validate_signed_native.py
```

The current six-suite discovery covers 68 tests: 19, 8, 7, 5, 3 and 26.
The existing reconciliation receipt remains external at
`(deleted-coord-artifact)/tooling-coverage-reconciliation-1c2ce1f7d5.json` (8,456 bytes,
SHA-256 `835eba64e99b0801d630eb75fd7a2da4ad3c1ab100517c9be8fa4112fd801a4d`).

## Combined-run repair

The first combined 68-test run exited 1 because
`tools/test_opc_read_ahead_timing_guard.py` used `marker.exists()` as the
handshake for a forked escaped child. File creation can become visible before
the PID bytes are written, producing `ValueError` from an empty marker. The
preserved pre-repair log is external:

```text
(deleted-coord-artifact)/tooling-coverage-1079-combined-before.log
1604 bytes
c78e4c3a5bceba8920c0feba72fe99328d06a6874c660b3f464f7081158e63c4
```

The test now uses a bounded two-second handshake that waits for a non-empty,
positive PID before refreshing ownership. It retains the escaped-child
identity assertion and the subsequent cleanup/reap assertions; it does not
skip, relax, or suppress the lifecycle check.

The exact combined command then passed all 68 tests in 19.812 seconds. The
committed repair tip was rerun unchanged; its external log is:

```text
(deleted-coord-artifact)/tooling-coverage-cc0d41a952-combined.log
901 bytes
f8aafbbd12d9ac2374dba8710df28ee57d2e08c4ec01959e875df0757101a3df
```

The run remains source-only tooling evidence. The normal safe-runner fixture
prints its intentional nonzero-child diagnostic and a `ResourceWarning`, but
the combined unittest result is `OK` with no failed or skipped tests.
