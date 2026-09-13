# Change 0552: compact XLSX source-cell offsets

Status: candidate rejected by the frozen main RSS, planning-guard, and cap admission gates. Baseline production has been restored, and the baseline-compatible public regression tests are retained. All eleven final quality checks passed. Decision, verification, and cleanup records document batch closure; no optimization is accepted in this batch.

OLE2 and OOXML performance have priority. ODF optimization is deferred until that goal completes; iWork is excluded from this work.

## Mechanism and scope

The candidate collects bounded cell offsets during the existing eligible worksheet validation/parser traversal. Source-backed `MultiSourceEdit::commit` can reuse those offsets to avoid reconstructing the complete worksheet layout. The proof is bound to the immutable worksheet source allocation and its parser Store. Unsupported actions, uncertain structure, mismatched ownership, or optional proof refusal retain the complete writer path. Proof-only refusal drops the optional collector while the authoritative parser continues. Rewritten output validation and independent readback remain required.

The retained proof uses 8-byte cell spans and a 2 MiB logical metadata budget with checked, fallible growth. Existing source eligibility remains bounded by 8 MiB and 131,072 provisional events. The logical metadata budget does not cover all transient parser/writer allocations or process RSS. The single-sheet `SourceEdit` API is outside the optimized and measured owner.

## Evidence and provenance

- `metrics-analysis.json` and `guards-analysis.json` validate the completed comparison and record failed admission gates. `analyzer-amendment.json` preserves a post-capture metadata-only timestamp serialization correction and the original failed guard-analyzer receipt; the gates and measured inputs did not change.
- `final-source-restoration.json` and `final-restoration-completion.json` bind the restored production and unchanged retained tests.
- `plan.json` freezes the scenarios, sample counts, ABBA order, admission gates, and scope limits before candidate measurement.
- `frozen-inputs.json`, `supplemental-inputs.json`, and `analysis-inputs.json` bind the drivers, ignored workspace lock, quality plan, and analyzers. The original baseline driver did not check the ignored workspace lock per child; the separately frozen lock and completion record establish its stated timing boundary. Later guarded captures check it per child.
- `candidate-attempts/` preserves source snapshots, patches, and failed-check provenance. Draft06 was reapplied cumulatively from the restored baseline after validating the public exact-output oracle. Draft07 changes only an immediate-return `while` to an equivalent `if` for Clippy.
- `candidate-source-binding.json` binds the frozen measurement manifest to draft07 and the public test files.
- `baseline-correctness.json` records historical baseline owner tests. `preflight-summary-draft07.json` distinguishes the draft06 16-proof/74-integration test results from draft07 feature/lint/format/boundary results. Final exact-source quality is still required.
- `public-exact-test-sources/` preserves the oracle correction chain. Publication of formatting whitespace between elements is an existing baseline refusal; the final tests cover that refusal and supported exact output separately.
- `proof-review-draft07.md` is the current source review. The preserved earlier reports are historical reviews, not additional approval of the final candidate.

## Serial measurement and decision workflow

Only the root coordinator runs builds, tests, and captures. Source and frozen scripts must remain unchanged throughout a capture. The owned target is `/home/zhuhe/litchi-goal-0552-target`; retained baseline binaries must remain available until the comparison and cleanup custody are complete.

The baseline first repeat is already retained. The subsequent commands are:

```sh
python3 -B docs/performance/results/change-0552/guarded_capture.py candidate
python3 -B docs/performance/results/change-0552/guarded_capture.py baseline-r2
```

The candidate command creates its stage once and captures both candidate repeats. It must not be restarted while its process is live or rerun over existing evidence. A terminal failure requires an explicit preserved-attempt decision. Baseline repeat two uses retained baseline binaries under the candidate execution-source manifest.

Main workflow and planning/cap analyzers require the complete matrix. Native latency and instrumented allocation measurements are separate; allocator elapsed time is not latency evidence. Every frozen mandatory gate must pass. Individual adverse and repeat-drift rows require review; aggregate averages cannot hide a failing row. Exact commit profiles are conditional on the native/memory pilot passing. Any additional large-tag/cell guard remains prospective until its inputs and actual measurements are recorded.

Acceptance also requires the full eleven-command quality plan on the final source, exact disposition/source custody, independent verification, documentation and seal, then commit and owned-target cleanup. If a mandatory gate fails, restore baseline production while retaining baseline-compatible public tests and the failed experiment evidence. Do not relax the gates to admit the candidate.
