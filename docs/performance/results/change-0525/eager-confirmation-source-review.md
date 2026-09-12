# Eager confirmation source review

This is a read-only audit of the frozen supplemental confirmation wrapper and
plan. No builds or captures were run.

| Artifact | SHA-256 |
| --- | --- |
| `eager_confirmation_guard.py` | `2485c9f70cdb475c1eed0bc25d266de3a8ab18739090d1d6c53ae4aaa2e36ca4` |
| `eager-confirmation-plan.json` | `08bd39c750ef27ef4a337b603f67e432f40f59c8c3a5e424a1d0758d388b0eb9` |
| `analyze_eager_guard.py` | `d4317d90bc523389c66d417fcbdeb6c330eee44f59862964192104e2d15cf220` |
| `run.py` | `1415937a8ee4602d5d8205c1b7fb98b84fd3d002d44758d3c5b8fe84a283b75d` |

## Blockers before final replay

1. **The post-cleanup replay cannot invoke the wrapper.** `verify.py` calls
   `run_replay` with `analyze --output <temporary-path>`, but
   `eager_confirmation_guard.py` accepts only `capture`, `validate`, or
   `analyze` and has no `--output` option. Its `analyze()` also always writes
   `eager-confirmation-comparison.json`. The final verifier therefore receives
   an argparse error, and direct analysis would overwrite retained evidence
   instead of producing stable output in the replay directory. The wrapper
   needs an explicit output destination that is used by replay.

2. **The frozen plan child schema does not match the root verifier.** Every
   child in `eager-confirmation-plan.json` includes
   `working_source_manifest`; `verify.py`'s `expected_children` dictionaries
   omit that key and are compared with exact Python dictionary equality. The
   plan validator will fail before child evidence is checked. The expected
   matrix must include the working-manifest binding.

3. **The report isolation flag is checked against the wrong concept.** The
   plan's `custody.fresh_child_per_sample: false` correctly describes the
   wrapper's one-subprocess-per-capture schedule. The harness report's
   `configuration.filesystem_fresh_child_per_sample` is a separate filesystem
   policy and is `true` in the retained oracle; the companion validator also
   requires `true`. `verify.py:1628` currently requires `false`, so a valid
   confirmation report will be rejected. That check must follow the report
   schema while retaining the plan-level capture custody check.

4. **Root replay does not consume the working-source binding.** The wrapper
   checks `working_source_manifest` against the live source before and after
   each child, but `verify.py` validates only `child["source_manifest"]` and
   ignores the working path and receipt/binding working hash. Once the child
   schema is corrected, the root verifier should validate both build and live
   working manifests. This is required for the baseline A1/A2 children, whose
   baseline binary is deliberately run against the candidate working source.

## Oracle and statistics review

The wrapper's `validate_raw()` is structurally sound: it delegates the complete
eager envelope to `analyze_eager_guard.py`, which calls canonical
`BASE.verify_elapsed` and `validate_sink`. It then binds corpus, full sink
values/buckets, and output digest to the frozen dense-sparse oracle. Pair
comparison retains p50, p95, p99, mean, peak RSS, and all over-five-percent
adverse rows. This satisfies the intended raw-statistics and sink-vector path
once replay can run.

The root verifier has a weaker duplicate validator. `validate_confirmation_elapsed`
only checks vector shape/permutation and accepts supplied summary statistics;
`validate_confirmation_raw` checks only sink status/write status and does not
reuse the companion's complete sink-vector checks. The later wrapper replay
would catch forged summaries after the output-destination blocker is fixed, but
the retained-artifact precheck should either delegate to the same companion
validator or independently recompute all elapsed statistics and sink vectors.

The ABBA capture state, serial receipt intervals, before/after source and binary
hashes, exact corpus/sink/output oracle, and post-cleanup binary custody checks
are otherwise aligned. No additional raw-capture blocker was found beyond the
replay, plan-schema, isolation-flag, and working-source-verification issues
above.

## Resolution before cleanup

The root verifier now redirects only the frozen wrapper’s comparison write to
a temporary destination, checks both build and working manifests, and keeps
the report filesystem policy separate from per-capture subprocess custody.
The wrapper and plan remain at the hashes above. Its replay performs canonical
statistics and complete sink-vector validation; the duplicate root precheck is
not the sole admission check. Root ran `validate_all` directly and reached only
the final missing cleanup receipt after all preceding checks passed; see
[`pre-cleanup-component-verification.json`](pre-cleanup-component-verification.json).
The original findings above are retained as the audit history.
