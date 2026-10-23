# 0362uv — Final gate execution readiness

The coordinator's safe runner still has 53 steps. Its `tooling-unit` step
now includes `tools.test_final_api_inventory`. The separate final inventory
must capture all 68 direct libraries with all features and hidden items;
the glossary's 55 profiles cover 41 packages and cannot replace it.

The runner now clears inherited `LITCHI_CAPPED` before invoking its Cargo
wrapper. A resource or provenance exception preserves existing child output
and writes a separate `.blocked.log`, with both hashes in the failure row.
It no longer overwrites useful command output when post-command checks fail.
This is execution readiness, not a completed gate or performance receipt.

## Runner identity and reproduction

The external runner is
`/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py`.
The preserved predecessor is `run_safe_gate-before-0362uv.py` beside it.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| Predecessor | 22,130 | `592ff5f7ad8493641ba6df7308bd66a97bddced35fc616c5803aacc5d30dcd2c` |
| Repaired runner | 22,721 | `d513002381419a3c09fe0457c644db1231eb8f8bb3094896525c8e8545f90c84` |
| `results/0362uv/safe-gate.patch` | 1,629 | `45a8b8fb8c8fdcc334d357e72b832330acfebe2ed6fe410b9b2905e65de93822` |

To reconstruct the repaired source, copy the preserved predecessor to a fresh
external directory as `run_safe_gate.py`, then apply the tracked patch there
with `patch -p1 --input <absolute-path-to-safe-gate.patch>`. The runner keeps
using its preserved sibling `safe_chunked_tests.py` and `bin/cargo`; a real
gate launch records all three source hashes in its plan.

The build-free regression command passes:

```text
python3 docs/unified-ops/results/0362uv/verify_safe_gate.py --runner /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/run_safe_gate.py --root /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration --scratch /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe/verification-tmp
```

It verifies the 53-step plan, inventory test registration, cap-marker
isolation, and preservation of exact child output after an injected
post-command resource failure, including both recorded log identities.

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

The coordinator must additionally verify that no gate-owned build children
remain and remeasure target size after the gate exits, before accepting the
final receipt or starting a measurement. The runner's per-step size sample
alone does not prove that a detached child has exited. A blocked retry may
retain prior child output beside its new failure log; that row is not a new
command pass. Use a fresh state directory when a distinct attempt receipt is
needed rather than attributing previous output to a new execution.
