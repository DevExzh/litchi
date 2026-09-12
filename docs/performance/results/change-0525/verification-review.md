# 0525 XLSX reconstruction evidence verification

`scope: independent staged verifier for the source-bound XLSX omission and
semantic-store reconstruction candidate`

`performance_claim: none while capture artifacts are incomplete`

This review owns the verification contract for change 0525. The frozen plan
is [`plan.json`](plan.json), SHA-256
`78d3fad228e4bd00476148044bbcffc8d48ec02f1e825b59c95385f92da40f97`; the
frozen capture driver is [`run.py`](run.py), SHA-256
`1415937a8ee4602d5d8205c1b7fb98b84fd3d002d44758d3c5b8fe84a283b75d`.
The baseline manifest currently retained in this bundle has SHA-256
`4838d157514eb9479d080d475f32921f97c30acc3ca59da43ceb1626cd9c5d03` and an
empty baseline patch. The prepared focused test patch is a draft, not frozen
candidate evidence; it is retained with SHA-256
`afaffd3571a018e8611ba2851b9f795ecef6788970106023138c794217b1c56f`.

The static closure translator and report are frozen at SHA-256
`bc92a6956a3349db42bf98e583ba16e8a4a41d07608ede7934f1ba5e37be71cd` and
`8ce112e24441e8d79ecc985a6139559320e998bdbf8aafb7671c572f99d7a7e9`.

The verifier first checks the accepted ADR hashes, the retained 0522 XLSX
source binding, the source-derived closure counts, and the frozen design and
adversarial reviews. It replays each captured source patch through a private
Git index and compares every replayed blob with its stage manifest. Candidate
differences must be nonempty and remain under `crates/litchi-xlsx/`; a
candidate-only new file requires a stage-local hash and content sidecar only
when complete patch replay does not reconstruct its blob.

For a completed campaign it checks every build and capture receipt, its
stage-local artifacts, the exact `TMPDIR`, binary and plan/script bindings,
the serial ABBA ordering, and the special retained baseline A2 binding. It
then reruns the frozen numeric, Callgrind, whole-child hardware, and eager
guard analyzers into temporary destinations and requires byte-for-byte report
replay. The numeric report supplies the four primary p50 gates; the profile
report supplies the four independent commit-Ir gates. Hardware counters stay
diagnostic and whole-child. Every adverse flag must have a retained review.
The eager capture authority remains `eager_guard.py`; its source-free report
schema is replayed through the companion `analyze_eager_guard.py` validator.
The standalone eager adverse review (or an explicitly embedded eager review
section) must bind `eager-guard-comparison.json`, retain every eager adverse
row and review string, record its non-applicable phase scope, and keep the
no-primary-gain boundary.
The supplemental dense-sparse confirmation is a separate frozen campaign:
`eager-confirmation-plan.json` binds `eager_confirmation_guard.py`, the four
children under `eager-confirmation/` must execute in A1/B1/B2/A2 order, and
each child must bind both its build source manifest and its retained working
source manifest. This preserves the baseline A1/A2 executions under the
candidate checkout. The plan's one-fresh-child-per-capture custody remains
separate from the raw report's filesystem-fresh-child-per-sample field, which
is validated against the companion report schema. The frozen wrapper has no
output option, so verification imports it in a subprocess, redirects only its
comparison write to a temporary path, and requires byte-for-byte replay while
leaving the retained comparison untouched. `eager-confirmation-comparison.json`
plus its adverse review must be retained before disposition; an accepted
disposition additionally requires the comparison's supplemental gate to be
true.

The quality summary is bound to all twelve frozen checks and its test counts
are recomputed from their logs. A decision or disposition record chooses the
final source manifest by hash; the verifier checks the live checkout against
that selected manifest and accepts either an adopted candidate or a rejected
candidate with an independently recorded final source. It does not infer the
disposition from timing results. Cleanup is accepted only with an exact owned
path list, no accessible process references, absent binaries/target trees,
and no Python caches.

The final decision record must explicitly contain `disposition`, `final_source`,
`final_source_manifest_sha256`, `production_change_retained`, and exact
`plan_sha256`, `comparison_sha256`, `profile_analysis_sha256`,
`quality_summary_sha256`, and `adverse_review_sha256` bindings. The quality
record has `status: pass`, twelve `{name, stage, receipt_sha256, exit_code,
executed_tests}` rows, and their aggregate `executed_tests`; a successful
`preflight-N` XLSX suite may supply its row only when that preflight manifest
hash equals the canonical candidate manifest. The cleanup record contains the
exact `plan_sha256`, `removed` owned paths, an empty
`accessible_process_references`, and true `owned_paths_absent` and
`python_cache_absent` fields.
The decision also binds `eager_comparison_sha256` and
`eager_adverse_review_sha256`, plus
`eager_confirmation_comparison_sha256` and
`eager_confirmation_adverse_review_sha256`.

The bounded negative-probe harness is [`verify_test.py`](verify_test.py). It
uses retained preflight-1 evidence and in-memory values to prove rejection of
a source-manifest mismatch, a non-equivalent preflight alias, an alias to a
failed preflight receipt, an inverted native ABBA block order, a duplicate
primary row, a cleanup record without the plan binding, and an unsafe quality
receipt path. Running it performs no build, capture, analyzer replay, source
edit, frozen-tool edit, or retained-artifact mutation; it prints a seven-check
JSON receipt and accepts an optional output path for callers that want to
retain that small probe result.

Until the candidate manifest, both complete capture stages, reports, quality
receipts, disposition, and cleanup receipt exist, [`verify.py`](verify.py)
returns `status: incomplete`. A partial or failed receipt is retained as
evidence and cannot be converted into a performance pass. OLE2/OOXML remains
the active priority; ODF is deferred and iWork is outside this review.
