# 0524 final source and evidence review

`scope: bounded read-only review of the frozen final source and retained 0524 evidence`

`decision: reject the production CFB candidate; retain independent tests and the existing test stabilization`

The final source restores the production implementation after the
`CheckedBitSet::test_and_set` candidate failed the native admission gate. This
review checked the final source patch, final manifest, candidate and final
receipts, numerical comparison, profile and hardware summaries, the filesystem
identity probe, and the retained test correction. It made no source, build,
test, capture, or commit changes.

## Final source custody

The frozen [`final/source-manifest.json`](final/source-manifest.json) has
SHA-256
`4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03` and
contains 8,584 source entries. Its relevant hashes are:

- `crates/litchi-cfb/src/file.rs`:
  `62f7b357d35b5bb2921336bb6a983a8d7e9ec58c30ae694a3be4c89e9c251788`
- `crates/litchi-cfb/src/writer/sequential.rs`:
  `177733997ab0b566424c35cf0cb1dc1a4b6182ed3e64c435ab1be93b940475ff`
- `tools/perf-baseline/src/lib.rs`:
  `1d4f9ed81ffed0a65d3ff49572c50f2e8bf60867d4befeffc77d2174526fbafd`

The final [`source.patch`](final/source.patch) has SHA-256
`1d49d642a640734abc00a6061337575ae4f23a61635e129ff819af943698bef0`. Its
only hunks are the independent
[`scratch_differential_matches_owned_chain_helper_and_resets`](../../../../crates/litchi-cfb/src/file.rs#L3240)
test and the existing
[`detected_temp_substitution_is_not_deleted_by_cleanup`](../../../../crates/litchi-cfb/src/writer/sequential.rs#L2089)
test stabilization. The production `CheckedBitSet` and collector code are
therefore restored; no candidate `test_and_set` method remains in the final
source.

The earlier [`source-review.md`](source-review.md) is the historical review of
the applied candidate source, whose `file.rs` hash was
`ec44f1209fe7b895bb3e2d55316eb3cf9e890017a74d6925fe9027e88370e737` under
candidate manifest SHA-256
`7720b9dc1a135e411def3b9164f60b37695fbd73a46c69fb455c3c145d4f84ce`.
Those candidate hashes and conclusions must not be read as the final
production source state.

## Candidate decision and evidence

The matched [`comparison.json`](comparison.json) and
[`adverse-review.json`](adverse-review.json) validate source, binary, corpus,
and output identity. The native primary gate is false: the four primary XLS
workflow p50 deltas range from `+1.46%` to `+4.87%` in repeat 1 and from
`-2.37%` to `+0.28%` in repeat 2. No workflow reaches the required 3% lower
p50 in both repeats. The CFB few-large guard is also slower by about
`+2.35%` to `+2.41%`. `adverse-review.json` records
`adoption_allowed: false` for this reason.

All 43 matched over-5% flags remain retained, including repeated whole-child
`rss.system_seconds` records; they are not treated as independent operation
regressions. All 60 same-build variation flags remain retained and reviewed.
The allocation guard passes: all 24 matched rows preserve identical
allocation-call, allocated-byte, and incremental-region-peak vectors. Absolute
live and whole-child measurements remain separate diagnostics.

The [`profile-comparison.json`](profile-comparison.json) status is `pass` with
valid stage reports and matched identities. The `xls-owned` owner `Ir`
decreases by `1.189%` in repeat 1 and `1.165%` in repeat 2, while the aggregate
`collect_exact` self `Ir` decrease is a mechanism diagnostic. The profile
report explicitly makes no native admission claim. Baseline and candidate
[`hardware-analysis.json`](baseline/hardware-analysis.json) and
[`hardware-analysis.json`](candidate/hardware-analysis.json) both validate
their whole-child captures; they do not provide an operation-local hardware
or speedup claim. These diagnostics cannot override the failed native gate.

## Test and quality custody

The candidate preflight passed 278 CFB library tests under the owned disk-backed
temporary directory, with the source unchanged. A later candidate full CFB
quality check retained 277 passes and one existing
`detected_temp_substitution_is_not_deleted_by_cleanup` failure. The
[`quality-adjustment.json`](quality-adjustment.json) records the failure and
the correction rationale. The filesystem-only
[`temp-identity-probe.json`](temp-identity-probe.json) observed inode reuse in
100/100 unlink/recreate trials and 0/100 reuse when the displaced file stayed
alive through rename. It supports the failure mechanism without claiming to
observe the failed test's inode retrospectively.

The final patch keeps the collector differential independent of the rejected
candidate API. It compares the unchanged exact-chain helper with reusable
scratch collection across valid, empty, cyclic, marker, index, and table-size
errors, checks exact results or error text, verifies reset-on-error, and checks
buffer reuse. The sequential substitution test now keeps the displaced staged
file alive, renames it, creates the replacement, checks distinct identities,
and cleans both files explicitly.

The final quality test run completed 4,374 executions: CFB 306 in each of the
two feature configurations (612 total), XLS 1,345, DOC 1,187, and PPT 1,230.
Workspace, Clippy, and rustdoc checks also passed. The final receipts are bound to
manifest
`4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03`. At independent review time, boundary and claim checks, cleanup, and sealing
were pending; the retained candidate failure receipt remains part of the
evidence history.

The batch remains within the active OLE2/OOXML priority; ODF is deferred and
iWork is excluded.

## Root completion evidence

Both remaining gates subsequently passed. `verification.json` records passing
post-cleanup replay of all three source stages, 59 serial intervals, fourteen
final quality gates, eight reports and comparisons, and 160 annotations.
Both owned build paths and Python caches are absent. The final disposition is
rejected under the native gate; sealing is verified separately after inventory
creation.
