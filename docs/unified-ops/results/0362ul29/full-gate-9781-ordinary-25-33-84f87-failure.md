# Ordinary gate diagnostic: 9781 rows 25–33

This is a reduced diagnostic record, not an acceptance receipt.  The run used
the immutable production source `9781cf6d87e4d261d1dd9670e144e2db127d869b`
and the reviewed 0362ul29 tooling.  The predecessor `9451` endpoint remains
unchanged and is not relabeled by this run.

## Invocation and binding

The endpoint was fresh before launch:

```
/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781-ordinary-25-33-84f87
```

The selected steps were indexes 25–33, in this order:

```
tests-ops, tests-facade, tests-core, clippy-leaves, clippy-ods,
tests-leaves, tests-ods-full, clippy-workspace,
tests-workspace-excluding-ods-ops
```

The exact command was:

```
python3 /home/zhuhe/code/litchi-worktrees/.coord/gate-repair-0362ul29/source/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py \
  --root /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781cf6d87/source \
  --target /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781-ordinary-25-33-84f87/target \
  --state /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781-ordinary-25-33-84f87 \
  --only tests-ops,tests-facade,tests-core,clippy-leaves,clippy-ods,tests-leaves,tests-ods-full,clippy-workspace,tests-workspace-excluding-ods-ops
```

The run was offline, locked, `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`,
and used the 20,000,000,000-byte hard target cap with the 19,000,000,000-byte
live guard.  The plan records these bindings and these tooling hashes:

```
run_safe_gate.py       9858715863b381695959ee9edde38b7cf985c80b06603130199ff369a0493553
safe_chunked_tests.py  3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f
bin/cargo              495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
```

## Result

Row 25 (`tests-ops`) ran 37 helper chunks.  Chunks 0000–0035 passed (36
chunks); chunk 0036 failed.  The helper stopped at that first failed chunk,
retained its artifacts, and did not run rows 26–33.  The helper checks receipt
is 184,816 bytes, SHA-256
`9c4aca1c4de01aa72555f70fb4fa95e2f335107b953723aeb5fa243486428e3f`; it
contains 37 rows: 36 passed and one failed.

The source/test failure is in `ops-tests/0036.log` (380,107 bytes, SHA-256
`812711fb05ec081b7bca99166fdcc29001994bf0e176d1f7981027f8208de01a`).
`pptx_cell_properties` passed all three tests.  The one failing test is
`pptx_batch_equivalence::corpus_batches_equal_their_items_one_at_a_time` at
`crates/litchi-ops/tests/pptx_batch_equivalence.rs:199:13`, for
`test-data/poi/test-data/xmldsign/hello-world-signed.pptx`, seed 1.  The batch
returned `Refused(Some(0), "DocumentProtected")`, while the item-by-item run
returned `Refused(Some(0), "PathNotFound")`.

Separately, after the helper failed, the outer runner produced a fail-closed
post-exit diagnostic: its owned-process set falsely included the outer runner
itself (`pid=1347889`, `pgrp=1347889`, `ppid=884590`).  No helper, Cargo, or
descendant remained.  This is a tooling self-attribution failure, distinct
from the source/test mismatch.  The blocked receipt is 1,022 bytes, SHA-256
`6bbfacb683155ea614be74695ac590605724a958107f0f6ed52b660c6a0aa172`, and
states:

```
RuntimeError("owned child remained after child exit: [{'pid': 1347889, ... run_safe_gate.py ...}]; cleanup={'reason': 'owned-child-timeout', 'signal': None, 'seen': [], 'remaining': []}")
```

## Preserved evidence

All paths below are outside Git and remain immutable:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `plan.json` | 24,954 | `76340f960e247f9d486c4857b87c8d7c9cab3e95cf5b8f60e770b06c5318545c` |
| `state.json` | 4,611 | `0e993765bfb12e4b1516fc57d090de6e977bd164d938f1e132f9d2cf4bc961c9` |
| `25-tests-ops-attempt-1792872768279741796-1347889.log` | 147,454 | `95bffe1a79b07d6b282169389d677fb50568a7ac752dc4b7e55b3b97723e814d` |
| `25-tests-ops-attempt-1792872768279741796-1347889.blocked.log` | 1,022 | `6bbfacb683155ea614be74695ac590605724a958107f0f6ed52b660c6a0aa172` |
| `25-tests-ops-attempt-1792872768279741796-1347889-cargo-targets.tsv` | 4,754 | `1da79f5076c4f5d21c45eed1119c48f5c7525e0cb7bc5f6e6891476bffbfd713` |
| `ops-tests/checks.json` | 184,816 | `9c4aca1c4de01aa72555f70fb4fa95e2f335107b953723aeb5fa243486428e3f` |
| `ops-tests/0036.log` | 380,107 | `812711fb05ec081b7bca99166fdcc29001994bf0e176d1f7981027f8208de01a` |
| `artifact-quarantine/manifest.jsonl` | 111,383 | `949c4681851ed97cf908ff2d9a058494f20495f51c7d164620c82f5520ffa16a` |

At terminal inspection, `target` measured 2,944,467,242 bytes and the
quarantine measured 15,967,702,835 bytes.  The terminal proof was collected
at 2026-10-24T21:16:12Z with:

```
ps -eo pid=,ppid=,pgid=,sid=,stat=,comm=,args= --no-headers |
  rg 'full-gate-9781-ordinary-25-33-84f87|1347889|1348143|1348187|2028613' |
  rg -v 'rg|bash'
```

It produced no rows.  No retry, blessing, source mutation, or evidence
renaming was performed.
