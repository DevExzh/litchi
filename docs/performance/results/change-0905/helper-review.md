# 0905 helper static review

Status: **PASS (static review only)**. This review covers the creation-only
`run_v2.py`, `analyze_v3.py`, `plan-template.json`, `README.md` and
`METHOD.md`. No helper, workload, build, profiler or reducer was executed for
this review. The root coordinator owns plan freezing and all subsequent
execution.

## Frozen scope represented by the template

The template fixes plan ID 0905, seed 905009, CPU 12, the four exact case IDs,
qualification 2/0, native six blocks of 100/10, allocator two repeats of 3/2,
the fixed child environment and the unchanged primary/review thresholds. It
uses `status: template-only` and explicit replacement markers so an unfinished
source or binary binding cannot be mistaken for a capture-ready plan.

The source contract has one candidate file,
`crates/xml-minifier/src/audit.rs`, and a plan-supplied positive
`source.file_count`. The analyzer requires the baseline and candidate
manifests to contain that exact count and set, then requires the manifest delta
to equal the one selected file. Both the current worktree and the isolated
candidate source root are checked against the candidate manifest at analysis
time; the capture driver checks the full census at stage boundaries and the
selected file around each child.

## Path and receipt review

The retained 0902 path correction is present in both helpers. Descriptor
readers accept a path string, a `Path` value or an object containing `path`;
relative paths are resolved against the packet directory and absolute paths
remain absolute. Internal calls pass `Path` values through the same conversion
before hashing, opening or importing, so a Python API caller cannot trigger the
old string-only failure. The template's placeholder status and descriptors are
not accepted by either runtime helper until the root coordinator replaces them.

Both roles use the same write-once command receipt shape. Each role requires
exit code zero, a nonempty equal `inputs_before`/`inputs_after` list and
`inputs_unchanged: true`. The baseline role binds the selected file to the
baseline manifest and the candidate role binds it to the candidate manifest.
The isolated source root may be reused sequentially between those builds, so a
baseline receipt is not compared with the final candidate hash. Binary
descriptors remain executable and are bound to their build receipts, plan,
helper and leaf receipts.

## Case and schedule review

The driver and reducer independently enumerate the same qualification, native
and allocator jobs. Native rows are paired by block and case, with alternating
role and case order. The reducer admits native reduction only after a passing
qualification marker and keeps allocator and RSS evidence diagnostic. It
retains the DOCX limitation that the unchanged control has paragraph/commit
identity but no XLSX source counters or output digest.

The one-edit XLSX source check no longer assumes a zero untouched-read count or
a fixed literal of three. It receives the validated corpus worksheet count and
requires `unselected_worksheet_read_calls == worksheet_count -
selected_worksheet_count` with positive bytes. The one-percent paths require
all worksheets selected and zero untouched reads. This distinction is checked
for native and allocator reports while the source/output/sink identity checks
remain shared.

## Statistical and safety review

The reduction gate uses paired p50 ratios, 10,000 bootstrap draws and seed
905009. Both primary estimates must be at most 0.98 with upper 95% limits
strictly below 1.0. p50/p95/p99 ratios above 1.05 and per-role/case p50
spreads above 1.20 remain flags. No helper retries a child or overwrites an
existing leaf or aggregate. Failure artifacts are retained, and this packet
does not authorize a production adoption or a performance claim.

## Current helper hashes

These hashes bind the reviewed files before later root plan/capture artifacts:

| File | SHA-256 |
| --- | --- |
| `run_v2.py` | `c1de78bf5ec162fc0dec90de54ac3d7e67fe1db1fcf85201b6f9ef124d8441dc` |
| `analyze_v3.py` | `9d38d1d9c02a21a334b57792852563825b611628a9127a04986be33e5575e551` |
| `plan-template.json` | `d7c4fb59beb543acd0aa0d793464c9c7fe798ff07e6410641dfb9c3f361e7071b` |
| `README.md` | `13b437e2aa9647f14f3ae4c1ac498de12879a9cbd4dfc8fd1ac4e744df753545` |
| `METHOD.md` | `a827688fb869948f51071553011ffcdf39a5ec1402066af37b9e70efa9aeb578` |

The root coordinator should recompute these bindings after any review-driven
edit and before freezing `plan.json`.
