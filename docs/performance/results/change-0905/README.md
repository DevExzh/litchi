# 0905 paired XML audit candidate protocol

This packet contains creation-only helpers for a fresh paired capture of the
bounded XML audit candidate. No workload, build, profiler, or reducer was run
while this packet was prepared. The production candidate and its tests are
owned by the root coordinator; this packet supplies the custody, schedule and
reduction protocol once the binaries and source manifests are frozen.

The candidate source delta is exactly one file,
`crates/xml-minifier/src/audit.rs`. The baseline is the current 0904 harness
source inventory. The root coordinator must provide the final baseline and
candidate manifests and set `source.file_count` to the manifest count before
copying `plan-template.json` to `plan.json`. The template deliberately contains
placeholders and has `status: template-only`; neither helper accepts it as a
capture plan. Every source, validator, binary and receipt descriptor in the
frozen plan must carry its actual path, byte count and SHA-256.

The fresh roles are expected at these packet-relative paths after the root
coordinator builds them:

```text
binaries/baseline-native
binaries/baseline-allocator
binaries/candidate-native
binaries/candidate-allocator
```

The exact paths may be replaced by absolute paths in the frozen plan. Baseline
and candidate build receipts use the same write-once command receipt schema:
both require exit code zero, a nonempty equal `inputs_before`/`inputs_after`
census, and `inputs_unchanged: true`. The receipt for each role binds the
selected source file to that role's manifest: the baseline receipt to the
baseline hash and the candidate receipt to the candidate hash. The isolated
source root may be reused sequentially between those builds, so the baseline
receipt is not incorrectly forced to match the final candidate hash.

The plan fixes four fresh-process cases in this order:

1. `xlsx_source_backed_cell_values_one_percent_edit_save`, medium (primary)
2. the same one-percent case, dense-sparse (primary)
3. `xlsx_source_backed_cell_values_one_edit_save`, medium (XLSX control)
4. `docx_semantic_one_edit_save`, medium (DOCX control)

Qualification runs eight children (two roles × four cases), two samples and
zero warmups per child. Native capture runs six alternating paired blocks,
100 samples and ten warmups per child: 48 children and 4,800 samples. Even
blocks use baseline/candidate and canonical case order; odd blocks reverse
both. Allocator capture is separate: two repeats, 16 children, three samples
and two warmups per child. Qualification and allocator observations do not
enter the native benefit estimate.

The adoption gate is unchanged: both primary paired p50 median ratios must be
at most 0.98 and each 10,000-draw paired-median bootstrap upper 95% bound must
be below 1.0. Bootstrap seed is 905009. Retain every review flag for a p50,
p95 or p99 ratio above 1.05 and every per-role/case p50 spread above 1.20.
Controls diagnose scope and do not substitute for either primary gate.

Run only after the root coordinator has frozen `plan.json`, source roots and
all four role binaries:

```sh
python3 -B docs/performance/results/change-0905/run_v2.py check
python3 -B docs/performance/results/change-0905/run_v2.py qualification
python3 -B docs/performance/results/change-0905/run_v2.py native
python3 -B docs/performance/results/change-0905/run_v2.py allocator
```

The capture driver is write-once and never retries a child. A failed command,
source drift, binary drift or missing report retains its stdout, stderr,
resource record and receipt, writes a validation failure, and stops the stage.
Do not delete or replace those artifacts. The driver checks the complete
source manifest at stage boundaries and the selected file around every child.
Relative plan paths are resolved against this packet directory; absolute paths
are preserved. The helper accepts path strings, `Path` values and descriptors
consistently, which keeps the corrected 0902 path-handling behavior when the
root invokes it through a Python API.

Review retained evidence without launching a workload:

```sh
python3 -B docs/performance/results/change-0905/analyze_v3.py qualification
python3 -B docs/performance/results/change-0905/analyze_v3.py native
python3 -B docs/performance/results/change-0905/analyze_v3.py allocator
python3 -B docs/performance/results/change-0905/analyze_v3.py all
```

Qualification must be sealed before native reduction. The reducer checks
semantic/corpus/sink identity across roles and stages, preserves the DOCX
limitation that its unchanged harness reports paragraph and commit identities
without XLSX source counters or an output digest, and treats RSS as descriptive
whole-process residency. It does not infer branch coverage or a causal speedup
from allocation, call-count or resource observations.

The one-edit XLSX validator checks its actual worksheet topology. It expects
compressed-copy reads for every unselected worksheet, deriving the required
count as `worksheet_count - selected_worksheet_count`; it does not assume a
hard-coded three-read result. The one-percent cases require all worksheets to
be selected and zero unselected reads. This keeps the control honest if the
fixed corpus topology changes while preserving the existing source/read
identity checks.

The root coordinator must record any failed preflight/build/capture attempt,
version helpers before execution, preserve all partial artifacts, and clean
only owned duplicate outputs or marked temporary build/source roots after
replay. This packet itself has created no worktrees and has run no experiment.
