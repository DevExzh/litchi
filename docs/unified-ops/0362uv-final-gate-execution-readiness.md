# 0362uv — Final gate execution readiness

The coordinator's safe runner still has 53 steps. Its `tooling-unit` step
now includes `tools.test_final_api_inventory`. The separate final inventory
must capture all 68 direct libraries with all features and hidden items;
the glossary's 55 profiles cover 41 packages and cannot replace it.

The runner now clears inherited `LITCHI_CAPPED` before invoking its Cargo
wrapper. Every child runs in a fresh process group; after its top-level exit,
the runner rejects owned builders, waits for owned descendants to disappear,
and samples the selected target until it is stable, including after the final
step. A resource or provenance exception preserves existing child output and
writes a separate attempt-specific `.blocked.log`, with both hashes in the
failure row. Retries receive unique current logs and target traces, so an
earlier attempt cannot be credited.
This is execution readiness, not a completed gate or performance receipt.

## Runner identity and reproduction

The external runner is
`/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py`.
The preserved predecessor is `run_safe_gate-before-0362uv-postexit.py` beside it.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| Predecessor | 22,721 | `d513002381419a3c09fe0457c644db1231eb8f8bb3094896525c8e8545f90c84` |
| Repaired runner | 26,767 | `2585788b5161c9d85a1a4d9481c378e8ca9bf64283f0ca8ff91bf05f7d741c27` |
| `results/0362uv/safe-gate-postexit.patch` | 7,482 | `657a87d0f3952ae473c261bb7d21e83e4fa87a79c916b430c98ddc48341b4dc5` |

To reconstruct the repaired source, copy the preserved predecessor to a fresh
external directory as `run_safe_gate.py`, then apply
`results/0362uv/safe-gate-postexit.patch` there with
`patch -p1 --input <absolute-path-to-safe-gate-postexit.patch>`. The runner keeps
using its preserved sibling `safe_chunked_tests.py` and `bin/cargo`; a real
gate launch records all three source hashes in its plan.

The build-free regression command passes:

```text
python3 docs/unified-ops/results/0362uv/verify_safe_gate.py --runner /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py --root /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration --scratch /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/verification-tmp
```

It verifies the 53-step plan, inventory test registration, cap-marker
isolation, a real forked target mutation after the top-level child exits, the
owned-builder rejection branch, exact child output after an injected
post-command resource failure, and two blocked retries with distinct current
log identities.

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
