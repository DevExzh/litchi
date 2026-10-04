# 0902 paired bounded XML-audit capture

This packet contains the capture driver for the 0902 candidate.  The driver
does not decide whether the candidate is faster and does not replace the
independent reducer.  It runs a frozen, paired schedule and retains custody
records for every child.  The root coordinator writes `plan.json` only after
the baseline and candidate native and allocator binaries have been built and
their source/build receipts have been reviewed.

The candidate has two production source files.  The baseline manifest is the
retained 0894 `source.json`; the candidate manifest must contain the same
10,771 paths and differ at exactly the two paths named by
`source.candidate_files`.  The current repository (`root_worktree`) is checked
against the candidate manifest at the start and end of every stage.  The two
candidate paths are checked around every child.  The required isolated
`source_root` is also checked at full stage boundaries.  A source or binary
mismatch aborts capture before a new child is started.

The frozen plan schema is `litchi.paired-audit-capture-0902.v1`:

```json
{
  "schema": "litchi.paired-audit-capture-0902.v1",
  "status": "frozen",
  "plan_id": "0902",
  "seed": 902009,
  "cpu": 12,
  "workload_cwd": "/home/zhuhe/code/litchi",
  "environment": {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
  "source": {
    "baseline_manifest": {"path": ".../change-0894/source.json", "bytes": 0, "sha256": "..."},
    "candidate_manifest": {"path": ".../change-0902/candidate-source.json", "bytes": 0, "sha256": "..."},
    "candidate_files": ["crates/.../first.rs", "crates/.../second.rs"],
    "root_worktree": "/home/zhuhe/code/litchi",
    "source_root": "/tmp/litchi-0902-src"
  },
  "semantic_validator": {"path": ".../change-0705/analyze.py", "bytes": 0, "sha256": "..."},
  "roles": {
    "baseline": {
      "native": {"path": ".../baseline-native", "bytes": 0, "sha256": "...", "build_receipt": {"path": ".../baseline-native-build.json", "bytes": 0, "sha256": "..."}},
      "allocator": {"path": ".../baseline-allocator", "bytes": 0, "sha256": "...", "build_receipt": {"path": ".../baseline-allocator-build.json", "bytes": 0, "sha256": "..."}}
    },
    "candidate": {
      "native": {"path": ".../candidate-native", "bytes": 0, "sha256": "...", "build_receipt": {"path": ".../candidate-native-run.json", "bytes": 0, "sha256": "..."}},
      "allocator": {"path": ".../candidate-allocator", "bytes": 0, "sha256": "...", "build_receipt": {"path": ".../candidate-allocator-run.json", "bytes": 0, "sha256": "..."}}
    }
  },
  "cases": [
    {"id": "primary-percent-medium", "kind": "primary", "case": "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape": "medium"},
    {"id": "primary-percent-dense-sparse", "kind": "primary", "case": "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape": "dense-sparse"},
    {"id": "control-xlsx-one-edit-medium", "kind": "control", "case": "xlsx_source_backed_cell_values_one_edit_save", "xlsx_cell_crud_shape": "medium"},
    {"id": "control-docx-one-edit-medium", "kind": "control", "case": "docx_semantic_one_edit_save", "semantic_shape": "medium"}
  ],
  "qualification": {"samples": 2, "warmup": 0},
  "native": {"blocks": 6, "samples": 100, "warmup": 10},
  "allocator": {"repeats": 2, "samples": 3, "warmup": 2},
  "resource": {"enabled": true, "time_path": "/usr/bin/time"},
  "gate": {
    "bootstrap": {"replicates": 10000, "seed": 902009, "statistic": "paired_median_ratio"},
    "primary": {"p50_median_ratio_max": 0.98, "upper_95_ci_lt": 1.0},
    "review": {"regression_ratio": 1.05, "p50_spread": 1.20, "tail_metrics": ["p50", "p95", "p99"]}
  }
}
```

Each binary and build receipt descriptor must include the actual `path`, byte
count, and SHA-256.  Baseline build receipts must have `status` pass and exit
code zero.  Candidate `command.py *-run.json` receipts use the frozen
input-census schema instead: they must have `exit_code: 0`,
`inputs_unchanged: true`, and nonempty equal `inputs_before` and `inputs_after`
arrays.  The reducer binds both candidate source files in those arrays to the
candidate manifest; a candidate receipt does not need a `status` field.

The native protocol first runs eight qualification children (four cases × two
roles), each with two samples and no warmup.  It then runs six blocks of eight
children (four cases × two roles), each with 100 samples and ten warmups.  Even
blocks use baseline/candidate order and the canonical case order above; odd
blocks use candidate/baseline order and the reversed case order.  This gives
48 native children and 4,800 primary/control samples.  Allocator capture is a
separate two-repeat schedule of 16 children (four cases × two roles × two
repeats), with three samples and two warmups per child.  No qualification or
allocator child is substituted for a native sample.

Run only after the root coordinator has frozen `plan.json` and reviewed the
binary/source receipts:

```sh
python3 -B docs/performance/results/change-0902/run.py check
python3 -B docs/performance/results/change-0902/run.py qualification
python3 -B docs/performance/results/change-0902/run.py native
python3 -B docs/performance/results/change-0902/run.py allocator
```

Review retained children without launching a workload.  Qualification must be
sealed first; native reduction then requires its write-once admission report.

```sh
python3 -B docs/performance/results/change-0902/analyze.py qualification
python3 -B docs/performance/results/change-0902/analyze.py native
python3 -B docs/performance/results/change-0902/analyze.py allocator
python3 -B docs/performance/results/change-0902/analyze.py all
```

The allocator review requires the allocator binary's
`system_allocator_operation_scoped` instrumentation and retains phase
allocation vectors and operation-metrics availability (`measured` or
`not_applicable`).  Allocator elapsed values remain diagnostics and never enter
the native acceptance gate.

`all` runs these stages in order.  A stage refuses to overwrite any existing
leaf, receipt, or aggregate marker.  A nonzero child, missing report/resource
file, source drift, or binary drift writes a retained `validation-failure.json`
and `receipt.json`, then stops without retrying.  Successful leaves contain
`start.json`, `command.json`, `report.json`, `stdout.txt`, `stderr.txt`,
`resource.json`, and `receipt.json` under `qualification/`, `native/`, or
`allocator/`.  `capture.json` is written only after every child in that stage
passes.  Resource values from `/usr/bin/time` are descriptive whole-process
observations; they do not establish a heap or latency claim.

The reducer must compare baseline and candidate reports by case, block, and
role, requiring equal non-timing corpus/output/sink/source-counter identities.
Primary acceptance remains the paired bootstrap rule in the experiment plan:
both primary p50 median ratios must be at most 0.98 and their upper 95% limits
must be below 1.0.  Controls and p50/p95/p99 regression flags remain review
signals, not substitutes for the primary gate.
