# Change 0899 evidence

Read-only attribution packet for [the report](../../0899-xml-validation-cost-attribution.md).
`analysis.json` is canonical; `review.json` is the independent numerical and
custody review. `source-audit.md` and `mce-start-audit.md` document source
constraints and deferred hypotheses. No candidate is adopted.

From the repository root, replay the reduction to a fresh output path:

```sh
python3 -B docs/performance/results/change-0899/reduce.py --output /tmp/0899-analysis-replay.json
cmp docs/performance/results/change-0899/analysis.json /tmp/0899-analysis-replay.json
```

Choose an absent output path and remove only the replay file you create.
`replay-cleanup.json` records the root's byte-identical replay and cleanup.
Annotation command receipts preserve exact offline `callgrind_annotate`
commands and inputs. The text outputs are immutable raw tool exports.

`allocation-custody-replay.json` records the strict 0895 replay.
`profile-content-replay.json` records the bounded 0892 content/semantic replay;
it does not claim strict historical mode identity. `replay-failures.json`
retains the strict mode mismatch and first adapter import failure.
`replay_profile.py` is the failed adapter; `replay_profile_v2.py` is its
successful successor. Their fixed output paths already exist, so do not
rerun them in place or overwrite historical evidence. `source-identity.json`
binds production/harness source identities and the quick-xml archive.
`final-review.json` is the root closure audit, not an independent final prose
review. No builds, tests, capture workloads, or new worktrees occurred.
