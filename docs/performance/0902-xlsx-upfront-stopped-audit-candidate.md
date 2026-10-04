# 0902 — reject the upfront-stopped XML audit observer bypass

The candidate is **rejected and restored**. Neither primary XLSX case meets the
frozen 2% end-to-end p50 benefit gate. Both complete XML audits were preserved,
and correctness checks passed, but correctness alone does not justify retaining
this specialization. Production and tests are back to their exact pre-candidate
bytes. The patch, candidate source overlay, binaries, raw captures and failed
attempts remain in [the evidence packet](results/change-0902/RESULT.md).

## Candidate and scope

Change 0901 attributed 58.57% / 61.88% of the measured publication owner's guest
instructions to XML auditing. That aggregate did not establish the cost or
coverage of the upfront-stopped window-search branch. This candidate called the
existing `verify_source` directly when the existing strict-half condition had
already disabled the observer. It still audited original then replacement with
the same policy, limits, errors and offsets. Identical input, usable windows,
exact-half differences and searches stopped later were unchanged.

The baseline binaries are retained 0894 builds of source `506aa2edae66`; the
10,771 selected files still match the pre-candidate tree at `23bf87661e`.
Candidate builds use the same source inventory with exactly two overrides,
Rust 1.95.0, release opt-level 3, fat LTO, one codegen unit and debug level 1.
Builds, source manifests, binaries, helpers and commands are hash-bound.
Isolated build roots lack Git metadata; workload PATH excludes rustc. Null
runtime Git/compiler fields are expected, with external custody authoritative.

All measurements are warm, serial fresh processes pinned to CPU 12. Eight
qualification children contain 16 samples. Six paired native blocks contain
48 children / 4,800 samples, 100 samples and ten warmups per child. Role order
alternates and case order reverses on odd blocks. Sixteen separate allocator
children contain 48 samples, three samples and two warmups per child. No
qualification or allocator timing enters the native benefit estimate.

## Native result and decision

The gate requires **both** primary paired p50 median ratios ≤0.98 and bootstrap
95% upper bounds <1, using 10,000 draws and seed 902009. Ratios below one favor
the candidate. The p50 columns below are medians of six child p50s; their ratio
is deliberately distinct from the paired-median statistic used for acceptance.

| Case | Baseline p50 median, ms | Candidate p50 median, ms | Paired p50 ratio | 95% interval |
| --- | ---: | ---: | ---: | --- |
| XLSX one-percent, medium | 16.624998 | 16.667102 | 1.003391 | 0.997852–1.006408 |
| XLSX one-percent, dense-sparse | 32.978559 | 32.923994 | 1.002641 | 0.986454–1.027383 |
| XLSX one-edit, medium control | 4.140779 | 4.114512 | 1.004749 | 0.988704–1.013353 |
| DOCX one-edit, medium control | 0.165631 | 0.157684 | 0.951525 | 0.938968–0.959989 |

No >1.05 p50/p95/p99 median-ratio review flag or >1.20 child-p50 spread flag
fires. The largest spread is 1.051900 for the candidate dense-sparse primary.
The DOCX control observation does not override the failed primary gate and is
not an adopted speedup. No workload row was retried, replaced or excluded.

Descriptively, publication phase p50 medians move from 7.356428 to 7.224574 ms
and 15.403525 to 15.162152 ms on the two primary shapes. Planning medians move
from 5.660360 to 5.807478 ms and 10.781533 to 11.107443 ms. These phase shifts
do not establish a causal mechanism or an end-to-end benefit; phase medians
cannot be added to reconstruct the median total. The total includes open,
planning, staged sets plus commit, and publication including the returned
snapshot drop. Reopen/oracles remain outside that total; the retained vectors
and descriptive summary expose their separate cost. DOCX's control timing
starts after package open and covers edit/commit/save.

## Allocation, RSS and preservation

All measured XLSX phase allocation, reallocation, deallocation and byte counters
match across roles and repeats. Publication observations are:

| XLSX case | Allocation calls | Allocated bytes | Region peak above entry live bytes |
| --- | ---: | ---: | ---: |
| One-percent, medium | 19,248 | 3,970,346 | 1,087,019 |
| One-percent, dense-sparse | 36,366 | 5,575,801 | 1,596,330 |
| One-edit, medium | 2,662 | 1,216,622 | 634,046 |

Absolute publication region peaks are three bytes higher in the candidate,
with the same peak above entry live bytes. Nested commit/core/staging regions
are not added as disjoint measurements. DOCX has no comparable phase allocation
regions; generic operation-metric availability is retained, not converted to
zero allocation claims.

Native process RSS medians (baseline → candidate, KiB) are 67,104 → 76,646,
84,196 → 93,806, 67,230 → 77,160 and 18,760 → 25,320 in table order above.
These increases remain visible. RSS includes executable residency and is not
an operation heap measurement; this batch makes no RSS improvement claim and
does not establish the cause of the difference.

Every serialized corpus, output, source-counter and sink identity matches
across roles and stages for XLSX. The one-edit control correctly records three
untouched compressed worksheet intersections totaling 20,330 bytes. Both
one-percent cases select all four worksheets and record zero unselected
worksheet intersections; this is not a selective-scaling claim. DOCX checks
reopened paragraph text and commit-operation counts in the unchanged harness,
but serializes no output digest or source counters. Its independent cross-role
identity guarantee is limited to corpus and sink evidence.

## Verification and retained failures

Final quality checks pass: 138 XML tests (one existing ignore), 6,296 Office
unit/integration tests (three existing ignores), 93 documentation tests (33
existing ignores), formatting, warning-denied Clippy, warning-denied rustdoc
and crate boundaries. Baseline XML tests pass 134 tests with one existing
ignore. The [ADR review](results/change-0902/adr-review.md) and independent
source/helper reviews preserve the original-first audit and window-proof
constraints. No new fuzz-coverage claim is made for this branch.

The following failures are retained rather than hidden:

- Initial candidate compilation rejected an unused `must_use` audit report;
  binding `_report` corrected it before the passing tests and release builds.
- Two isolated Office-test builds lacked existing fixture paths, first
  `3rdparty`, then `docs`. Read-only fixture links enabled the passing suite.
- The first capture helper rejected its own `Path` descriptor before launching
  any workload. Versioned `run_v2.py` and `analyze_v2.py` fix that type handling.
- Qualification analysis initially assumed zero unselected reads for the
  one-edit control. `analyze_v3.py` checks its three compressed-copy reads and
  replays the same successful qualification reports before native capture.
- The first descriptive-summary helper compared iteration-order phase totals
  with sorted latency samples. Its versioned correction applies `sample_order`;
  the acceptance analysis and captured observations are unchanged.

The canonical analysis admits all 72 children, with exact source/custody and
semantic replay. Independent raw review reproduces the paired ratios and
intervals exactly and checks 60 measured XLSX allocation vectors / 180 phase
samples; four DOCX allocator reports have no phase vectors. Historical replay also verifies all 10,771 baseline archive
files and both retained candidate overrides, and reproduces the raw reduction
without requiring the temporary checkout. `decision.json` records restoration
of exactly the two candidate files to HEAD/baseline hashes. No unrelated file
or worktree is part of this change.

## Next work

This disproves adoption of this candidate under the named gate, not all possible
observer costs or XML-audit optimizations. Before another specialization,
attribute actual audit branches and hot-loop costs rather than treating the
large aggregate audit share as removable work. Both mandatory validations and
all accepted ADR constraints remain in force. The broader non-iWork Office
performance goal remains active; ODF remains deferred and iWork excluded.
