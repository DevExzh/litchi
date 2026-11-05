# 0362uv — Final gate execution readiness

The coordinator's safe runner still has 53 steps. Its `tooling-unit` step
now includes `tools.test_final_api_inventory`. The separate final inventory
must capture all 68 direct libraries with all features and hidden items;
the glossary's 55 profiles cover 41 packages and cannot replace it.

The runner now clears inherited `LITCHI_CAPPED` before invoking its Cargo
wrapper. Every child runs in a fresh process group. On Linux the runner also
becomes a child subreaper and captures PID/start-time identities while the
top-level child is alive, so a fast fork that calls `setsid` remains owned
after reparenting. After top-level exit, the runner rejects owned builders,
waits for owned descendants to disappear, and samples the selected target
until it is stable, including after the final step. A resource or provenance
exception preserves existing child output and writes a separate
attempt-specific `.blocked.log`, with both hashes in the failure row. Retries
receive unique current logs and target traces, so an earlier attempt cannot be
credited. The `/proc` scanner skips only a PID-vanished `ProcessLookupError`
race; other inspection errors remain failures.
This is execution readiness, not a completed gate or performance receipt.

## Runner identity and reproduction

The external runner is
`/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py`.
The preserved predecessor for the current successor is
`run_safe_gate-before-0362uv-procrace.py` beside it. Earlier post-exit
successors and their predecessors remain beside them for provenance.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| Current predecessor | 32,771 | `8fedc488825e893f4101f855e9305775e8e9641aa405c790e76b75920ffce9c1` |
| Current repaired runner | 32,791 | `0691c5cb469d9f59caf67c63b748e4d2c80ea5b58e9b99cbde0e171fa5e8a372` |
| `results/0362uv/safe-gate-procrace.patch` | 236 | `971cf7ec3ec7e0076694351022a1218bff68ccd7df738beebd23747ce4deaa9e` |
| Prior pathless-adoption patch | 1,503 | `d238e9df0e0914525c4775e32bee9d771131a73734991684dc1db4f6165034d0` |
| Prior detached-tracking patch | 9,571 | `895e7331634dccf2b0d69fe2e923e2c7ebf4caf3bc85ab8b1f3c30013c3e4682` |
| Earlier predecessor | 22,721 | `d513002381419a3c09fe0457c644db1231eb8f8bb3094896525c8e8545f90c84` |
| `results/0362uv/safe-gate-postexit.patch` | 7,482 | `657a87d0f3952ae473c261bb7d21e83e4fa87a79c916b430c98ddc48341b4dc5` |

To reconstruct the current source, copy the current predecessor to a fresh
external directory as `run_safe_gate.py`, then apply
`results/0362uv/safe-gate-procrace.patch` there with
`patch -p1 --input <absolute-path-to-safe-gate-procrace.patch>`. The runner keeps
using its preserved sibling `safe_chunked_tests.py` and `bin/cargo`; a real
gate launch records all three source hashes in its plan.

The build-free regression command passes:

```text
python3 docs/unified-ops/results/0362uv/verify_safe_gate.py --runner /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py --root /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration --scratch /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/verification-tmp-procrace
```

It verifies the 53-step plan, inventory test registration, cap-marker
isolation, a real forked-and-`setsid` target mutation through an inherited
target file descriptor after `chdir("/")` with no target path in argv, the
pathless adopted-child branch, pre-existing direct-child lineage exclusion,
the `/proc` vanished-PID race, the owned-builder rejection branch, exact child output after an injected
post-command resource failure, and two blocked retries with distinct current
log identities. The predecessor fails this detached-helper assertion; the
current runner passes it.

## Remaining phases

Run the following build-free qualification tooling suites separately from
the 53-step source gate:

```text
python3 -m unittest tools.test_native_qualification_runner tools.test_odp_native_qualification tools.test_w8_file_provenance tools.test_w8_original_baseline tools.test_xls_f3_qualification_abba tools.test_xlsx_source_family_abba_batched
python3 -m unittest discover -s docs/unified-ops/results/0361fda -p test_iwa_native_qualification.py
```

The actual native/ABBA runs remain separate required receipts, including the
IWA constructor/owner, W4b/W5, ODP, W8, XLS F3, XLSX source families and the
new ODF shared-triad lane. No unit-test pass substitutes for those runs.

Run final inventory separately with `.coord/bin` first on `PATH`; its shim
verification expects that capped wrapper, not the gate-local wrapper. Both
phases use external output and temporary storage and may reuse one bounded
target cache, sequentially. Each still requires the 19.2-billion-byte target
preflight and 20-billion-byte hard ceiling. Isolated documentation steps
must start without unrelated builders; passing them is not a quiet-host
native timing claim. After an interrupted gate, verify its children have
exited before resuming or reusing the target.

The separate tooling commands passed 63 native-harness tests and 20 iWork
qualification tests. `results/0362uv/tooling-checks.json` records exact
commands and external log identities. The six-suite run straddled the
documentation-only `947b98f94f` commit; all tested tool sources are unchanged
from `a2f0984719`. The iWork suite ran on the clean successor. Independent
review confirms the runner patch and its regression proof.

The final gate and native/corpus receipts remain separate acceptance phases;
this record does not claim that any of them has passed. The repaired runner's
post-exit guard is part of every step, including the final one, and its
attempt-specific receipt paths are the identity used when resuming or
retrying.
