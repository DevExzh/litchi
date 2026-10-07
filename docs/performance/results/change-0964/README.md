# 0964 retained evidence

The [change record](../../0964-xlsx-direct-event-transport-trial.md) describes a
rejected direct-event-transport screen. Production is unchanged. Only epoch2
feeds [screen-v1.json](analysis-epoch2/screen-v1.json) and
[selection-v1.json](analysis-epoch2/selection-v1.json). Independent raw audit
[passes](review/independent-epoch2-raw-audit-v1.json) with zero ratio/selector
mismatches. The candidate patch is retained for inspection, not applied.

## Retention and size policy

Canonical external root: `/home/zhuhe/litchi-performance-0964`.
[external-evidence-inventory-v3.json](external-evidence-inventory-v3.json)
records each retained evidence file's relative name, bytes, mode and SHA-256.
Source trees are covered by their complete external manifests. Build caches
were deleted; raw captures, rejected epoch1, failed attempts, source snapshots,
and retained executables remain outside the checkout. No compiled executable
or evidence file of 1 MiB or more is copied into this directory.

| External artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `source-manifest-baseline-v2.json` | 2,267,944 | `74d3ead27a1518093be3f80fbe50cfa1280526e08957b1774dbc6fd891bc4eda` |
| `source-manifest-candidate-v4.json` | 2,268,238 | `19e4bdf18aee484ade90067a98ec6ab0ff5f6d017105840df4030f697f25779b` |
| `executables/xlsx_managed_compression_benchmark.v3` | 43,270,848 | `d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910` |
| `executables/xlsx_managed_compression_alloc.v3` | 43,688,016 | `3e722898ed854f6016e11cf53ef7fa158e3c16b2c9fd7b37fe09c4daea63d36d` |

The standalone harness lock is pinned in the build receipt. The baseline uses
0957's retained ordinary binaries, whose runtime/harness source content and
modes match the current export; a different Git revision alone is not evidence
of stale source. The root Cargo.lock overlay was used for quality checks only.
The exact build command, compiler environment, retained binary hashes and
terminal-before-binary parsing are in
[the build result](receipts/candidate-managed-v3-build-result.json).

## Reproduction

Commands below use retained absolute custody paths. Outputs are write-once;
choose new output paths, preserving old evidence. Recreating binaries or source
metadata on another host may change byte hashes; report the new identities.
The committed scripts retain the original absolute bindings intentionally.

```sh
R=/home/zhuhe/litchi-performance-0964
PYTHONDONTWRITEBYTECODE=1 python3 "$R/tools/reduce_0964_v1.py"   --plan "$R/protocol/plan-final-epoch2-v1.json"   --native-packet "$R/capture-native-epoch2-v1"   --allocator-packet "$R/capture-allocator-epoch2-v1"   --output "$R/analysis-epoch2/screen-replay.json"
PYTHONDONTWRITEBYTECODE=1 python3 "$R/tools/select_screen_0964_v1.py"   --report "$R/analysis-epoch2/screen-replay.json"   --output "$R/analysis-epoch2/selection-replay.json"
PYTHONDONTWRITEBYTECODE=1 python3 "$R/tools/audit_screen_v2.py" "$R" --output "$R/review/raw-audit-replay.json"
```

Source export regeneration starts with the exact `git archive` command in
[source-export-v1.json](receipts/source-export-v1.json), at commit
`93d32552d98586f42822b2dc9bbec7c8ef6cafc9`. Restore the root lock overlay and
recorded modes, then apply [candidate-v4.patch](candidate-v4.patch) only to a
separate candidate export. The full manifest census can be regenerated with:

```sh
python3 "$R/tools/freeze_candidate_scope_v1.py"   --baseline-source "$R/source-baseline"   --candidate-source "$R/source-candidate"   --baseline-manifest "$R/source-manifest-baseline-v2.json"   --candidate-template "$R/source-manifest-candidate-v4.json"   --test-allowlist-file "$R/protocol/candidate-test-allowlist-v3.json"   --output-manifest /tmp/0964-manifest-replay.json   --output-receipt /tmp/0964-freeze-replay.json
```

For an independent baseline census, hash each sorted file of the baseline
export by relative path, size, mode and SHA-256; compare against the baseline
manifest's `files` rows. Timestamps/path metadata are receipt metadata, not
runtime source identity. The quality receipt binds all exact test commands and
logs, including failed attempts and temporary fixture cleanup.

`tools/build_candidate_managed_v3.py` records the complete release build; use
fresh output locations if rebuilding. The diagnostic CPU findings replay the
frozen 0962 raw streams with `tools/analyze_source_selected_v1.py`, and compiled
call edges are checked by `tools/check_codegen_v1.py`. Their command/stream
bindings and the independent CPU review are included here.

Fresh workloads require a newly frozen plan and root coordinator GO receipt.
The retained epoch2 invocation was `screen_0964_v2.py --capture --plan
.../protocol/plan-final-epoch2-v1.json --lane native --go --packet
.../capture-native-epoch2-v1`, followed serially by the allocator lane, with
`LITCHI_PERF_0964_CAPTURE_APPROVED=1`. The plan binds full argv, environment,
affinity, order, samples, oracles and deadlines. Do not reuse raw output roots.

Epoch1's 54 workload-valid children are wholly excluded because every recorded
launch-affinity witness was the wrapper's pre-taskset mask. The unchanged
reducer refused them before ratios were computed. Epoch2 sets affinity in the
child before exec, verifies it before waiting, and binds positive/negative
controls. It is a separate full epoch, with no child replacement. The first
independent audit attempt also had a report-path typo; its failure receipt is
retained, and the corrected audit passes without changing data or gates.
