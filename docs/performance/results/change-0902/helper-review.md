# 0902 helper review

Status: **PASS (static helper review)**. This review covers the final
`run.py`, `analyze.py`, and the helper protocol described by `README.md`. It
reads the helpers, the frozen method/template, the retained 0705 validators,
and the current operation-metrics schema. It did not execute either helper or
launch a workload. Root owns plan freezing, builds, captures, tests, and the
eventual qualification/native decisions.

## Final helper bindings

The reviewed packet files are bound by these SHA-256 values:

* `run.py`:
  `3418daf9a5f83e166a31846c3d94b96769a218db64b206fac7f484c8b463fcd7`
* `analyze.py`:
  `d035220b5a3ae5a85c06eaa213d0fe5549002d068334c94d4cf0e770aadf24c5`
* `README.md`:
  `3d02d2f5ebe66ac64d19e9e0d7b685a69014b2fca42fd6f7d4f8d7144d0ec56b`

The final helper revisions include the retained compile-only quality receipt
correction in the packet documentation. This review does not treat that
receipt as workload evidence.

## Capture helper and custody

`run.py` pins the 0902 schema, plan identity, seed, gate constants, fixed child
environment, CPU command, binary role, case arguments, sample counts, warmups,
and alternating native schedule. It makes every child a fresh process, retains
stdout/stderr/report/resource/receipt artifacts, refuses retries and existing
outputs, checks the source census at stage boundaries and the candidate files
around every child, and aborts a stage on source, binary, command, or child
failure. Candidate build receipts are checked against the isolated
`source_root` input census. The command record's `source` field is correctly
paired with the receipt's `source_start`; the end and error fields are also
paired.

`analyze.py` independently rechecks the exact 10,771-file census and two-file
delta in both source roots, binary/build descriptors, candidate input hashes,
command and environment hashes, artifact descriptors including mode and
executable bits, retained receipts, and the exact expected child-folder set.
The final aggregate receipt list is compared with each retained leaf receipt.
These checks make the source, helper, binary, command, and report bindings
reviewable before reduction.

## Case and semantic validation

The analyzer accepts exactly these ordered IDs and classifications:

* `primary-percent-medium` and `primary-percent-dense-sparse`: the
  source-backed XLSX one-percent case, primary;
* `control-xlsx-one-edit-medium`: the source-backed XLSX one-edit case,
  control; and
* `control-docx-one-edit-medium`: the semantic DOCX one-edit case, control.

The case selector, shape selector, report configuration, result case,
sample/warmup counts, and role/stage schedule are bound to those IDs. The
XLSX primary uses the retained 0705 corpus, sink, statistics, source-vector,
phase-alignment, and unmanaged-cache validators. The one-edit XLSX control
has its own one-edit corpus and source checks, so it cannot be relabeled as a
one-percent corpus. Source/output digests, cache/read counters, phase vectors,
sink limits, and cross-role semantic identities are retained and compared.

Allocator reports require the operation-scoped instrumentation and pinned
`serialized_region_peak_v3` revision. Every phase allocation vector has the
expected sample cardinality, scope, status, nonnegative integer fields, and
peak/live boundary relations. Failed allocation counts are rejected. The
operation-metrics envelope requires aligned sample metadata and preserves
availability status; allocator timing and allocation observations are
diagnostics and do not enter the native acceptance gate.

The DOCX control deliberately requires its fixed semantic corpus and sink
invariants and rejects fabricated XLSX source/output fields. The workload's
unchanged DOCX harness reopens each output and checks paragraph text and
commit-operation counts. The report has no DOCX output digest or source
counter identity, so this reducer makes no independent DOCX byte-identity
claim, in accordance with `METHOD.md`.

Qualification is an explicit independent analyzer action. It emits a
write-once `qualification-analysis.json` with `status: pass` and
`qualification_admitted: true`; native and `all` reduction require that marker
to exist. This gives the coordinator a clear qualification-before-native gate
and keeps qualification rows out of native timing reduction.

## Gate statistics

`native_reduce` requires 48 native rows, pairs baseline and candidate by case
and block, and computes each block ratio from candidate p50 divided by
baseline p50. For each case it uses the frozen paired-median bootstrap with
10,000 draws and seed 902009. The retained lower and upper limits use the
nearest-rank indices 249 and 9,749 after sorting the draws. Only the two
primary cases can be accepted, and each must have median ratio at most 0.98
and upper 95% limit strictly below 1.0. Controls remain visible but cannot
pass the primary gate.

The review ledger keeps p50, p95, and p99 candidate/baseline median ratios for
each case and flags values above 1.05. It also flags per-role/case p50 spread
above 1.20 across the six native blocks. No qualification or allocator row is
used in these acceptance statistics, and RSS remains descriptive whole-process
evidence.

## Limitations and residual hardening notes

The helpers intentionally trust the frozen plan and retained workload schema
after validating the fields needed for this packet. Defensive improvements
that would be useful for a reusable hostile-input parser include rejecting
duplicate entries in the two-element `candidate_files` list, pinning the plan
CPU to 12 inside both helpers, validating every optional operation-metrics key
against the full generic schema, and checking all leaf stage/index/role/case
metadata fields independently of the command line. The frozen plan and exact
case schedule already provide these values for this run; none changes the
review verdict for the 0902 capture.

No performance, allocator, RSS, or production-adoption claim follows from
this static review. Those claims remain subject to root's frozen plan,
qualification result, native capture, and retained reduction evidence.
