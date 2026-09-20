# 0724 XLS checkpoint attribution packet

Diagnostic only; no production retention or performance/support claim. See the
[report](../../0724-xls-checkpoint-cost-attribution.md). The 0723 rejection remains.

`plan.json`, `capture.py`, `analyze.py`, fixture/probe/source censuses and eight
binary identities were frozen in `freeze.json` before all 480 captures.
`build.py` created baseline/layout/selection/full source archives, built two
unchanged probes per variant serially, and restored baseline bytes in `finally`.
`source-guard.py` independently proves the exact contrasts and restoration.

The independent terminal `audit.py` was finalized after capture. It and
`source-guard.py` are sealed by the final artifact manifest, not falsely claimed
as capture-frozen tools. The audit checks exact filenames and containment,
commands, order, hashes, metadata and outcomes, and independently recomputes
all raw stage statistics. `audit.json` and `analysis.json` stage maps match
exactly. `negative-checks.py` exercises the actual frozen primary analyzer on
temporary copies. No capture is modified by verifier controls.

Read-only replay from a checkout with the restored baseline source:

```text
python3 -B docs/performance/results/change-0724/source-guard.py
python3 -B docs/performance/results/change-0724/artifact-seal.py --check
```

Use a copy of this packet to run commands that regenerate analysis artifacts:

```text
python3 -B docs/performance/results/change-0724/analyze.py
python3 -B docs/performance/results/change-0724/audit.py
python3 -B docs/performance/results/change-0724/negative-checks.py
```

All return zero for valid diagnostic evidence; this is not a retention PASS.
The primary analyzer verifies frozen inputs and raw outcomes/statistics. The
independent audit also verifies live binaries or exact `cleanup.json` witnesses
when binaries are absent. Both must run for the complete evidence check.

Fresh capture requires a separate checkout of `0d943df447`, matching fixtures,
rebuilding the variants, adapting explicit target/binary paths consistently,
and fresh output directories. Run `build.py`, `source-guard.py`, then
`capture.py freeze` and `capture.py run`. Existing freeze/capture files are
never overwritten. Do not replace retained observations with a reproduction.

`initial-build-failure/` and `audit-schema-correction/` preserve tooling failures.
The only unused-code allowances are on intentionally unused diagnostic items.
No production lint is weakened. Raw logs keep their original bytes. The two
external owned roots are `/home/zhuhe/code/litchi-target-0724` and
`/home/zhuhe/code/litchi-0724-bin`; exact cleanup witnesses support final replay.
