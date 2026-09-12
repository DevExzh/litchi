# Root smoke review

The bounded scaffold review approves smoke wiring and evidence integrity only.
The retained run contains 21 lanes, one process and one measured sample per
lane. The repeated full profile has not run and remains gated. These samples
do not establish throughput, scaling, or a before/after performance improvement.

The independent review replayed the current verifier, checked all 25 manifest
inputs, and passed the five verifier regressions. Typed refusal paths exercise
the production API and verify physical republish, metadata, and reopened
semantics. Cleanup simulation retained full-mode and unrelated evidence.
Allocation metrics and whole-process RSS remain separate.

After measurement, README invocation examples were corrected from `sh` to
`bash`, matching the runners' Bash syntax. The manifest and provenance receipts
were refreshed for that prose-only change. Production source, harness source,
and recorded binary hashes did not change; measurements were not rerun for the
documentation correction. This review note records that distinction and is
not a measured input.

Root separately ran all five verifier regressions successfully before commit.
The owner reports its staging checkout and target were cleaned. No production
files are included in this evidence batch; `baseline-source/` is a retained
copy for replay, bound to the committed baseline named in the manifest.
