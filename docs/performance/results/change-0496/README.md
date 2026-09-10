# DOCX opened-edit phase diagnostics (0496)

This bundle measures the same opt-in diagnostic harness against production
revisions `de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8` (before) and
`44a4710699ef17041d5969240c30984dffbc3319` (after). It changes no production
code. See `plan.md` for the frozen hypothesis and scope, `protocol.json` for
the exact matrix, `builds.json` for the four retained executables, and
`analysis/formal1.json` for individual results and all review flags.

The phase clocks measure wall time, including instrumentation overhead; they
are not CPU profiles. Allocation counters cover one full lifecycle. GNU time
RSS covers the entire child, including corpus setup, preflight, verification,
and report generation. Normal binaries have no allocation counters. Managed
rows are after-only capability observations. The two reversed repeat blocks
provide descriptive uncertainty, not independent-host statistical confidence.
Earlier 0495 regression causes remain unresolved.

All 32 formal children and 960 samples passed. The independent scalar review
confirms publication is the largest named wall-time interval in all 16 normal
children. This interval includes copying output into the preallocated retaining sink;
output hashing and semantic verification occur after timing. It does not
isolate production CPU work. The analysis retains 74 flagged comparison rows:
60 phase-percentile rows, nine reallocation-count rows, four RSS rows, and one
full-lifecycle latency row. These are overlapping metric review triggers, not
74 independent regressed scenarios. Reallocation calls rise from 185 to 228
while total allocation calls and bytes fall in the unmanaged comparator.
See the per-change record for all scoped results.

The full release suite passed 479 tests (one ignored); Clippy, rustdoc,
formatting, boundary checks, and all 14 final capture-helper tests passed.
`helper-tests.json` is the historical 12-test attempt;
`helper-tests-final2.json` binds the frozen 14-test helper. `cleanup.json`
records 10.709 GiB of disposable source/build removal and preservation of
the four measured executables. Eight focused seal-helper tests and the
report-classification/CRUD-index documentation checks also passed; see
`final-helper-doc-gates.json` and `seal-helper-tests-final2.json`.
The initial seal attempt reached manifest creation after evidence replay but
failed on a missing timestamp helper; `development/seal-attempt1.json` retains
that failure, and the final helper tests exercise the corrected manifest path.

## Build and source reproduction

`before-harness.patch` and `after-harness.patch` apply to their corresponding
production revisions. The before patch includes the 0495 harness overlay.
`harness-source.json` binds the four identical harness files; the two
`builds/*-source.json` manifests bind all selected compilation inputs. Preserve
the recorded tracked harness lockfile and use `workspace-Cargo.lock` for the
workspace lockfile. Restore fixtures from the corresponding revisions.

Create separate checkouts at the recorded revisions, apply the corresponding
patch with `git apply --index`, and install the recorded workspace lockfile.
The exact Cargo commands and environment are in each build receipt under
`builds/`. Use Rust 1.98.1, release debug information, frame pointers, four build
jobs, and disabled incremental compilation as recorded. Build normal and
allocator executables separately and retain each executable before the next
build overwrites the shared output path. Hashes in `builds.json` identify the
actual measured executables; a rebuilt executable is a new experiment unless
its identity and all inputs match.

`build.py` is retained as the historical driver, including its mistaken debug
profile for the after tests. That debug test attempt was intentionally
terminated and is not a passing gate. Its partial logs and interruption
receipt remain under `builds/` and `development/`. `gates.py` contains the
accepted full release-profile test, Clippy, and rustdoc commands; the final2
receipts prove those gates passed without source changes. Do not use the
historical driver's debug gate as the accepted reproduction command.

## Capture and validation

From the repository root, the recorded capture commands are:

```sh
python3 -B docs/performance/results/change-0496/capture.py freeze
python3 -B docs/performance/results/change-0496/capture.py capture --attempt formal1
python3 -B docs/performance/results/change-0496/capture.py analyze --attempt formal1
python3 -B docs/performance/results/change-0496/capture.py verify --attempt formal1
```

These commands create immutable outputs and refuse to overwrite an existing
attempt. A fresh experiment needs fresh evidence paths and a newly bound
protocol. Each child receipt records its actual command, environment, source,
binary, start/finish times, terminal status, and private-directory cleanup.
The entire matrix holds the existing CPU measurement lock and pins children
to CPU 2. Original reports are retained unchanged. The phase validator checks
interval conservation and strips only the new row field into a temporary
projection for the unchanged sealed 0495 oracle validator.

After finalization, use the read-only verifier:

```sh
python3 -B docs/performance/results/change-0496/seal.py verify
```

The source checkouts and Cargo target are disposable; retained executables,
source patches/manifests, raw reports, terminal receipts, analysis, reviews,
and validation evidence remain. `cleanup.json` records final removal.
