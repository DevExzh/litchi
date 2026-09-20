# 0713 shared MCE substring-search evidence

The measured candidate replaces the private scalar byte-window search with the
existing `memchr::memmem::find` dependency. The 32-child pilot passes all 48
hard gates and exact deterministic output parity. Final disposition, quality
and resource review are recorded in the [change report](../../0713-mce-substring-search.md).

- `original-codec.rs`, `baseline-codec.rs`, `candidate-codec.rs`, both patches,
  and `source-preparation.json` distinguish original HEAD, the baseline with
  two cfg(test) guards, and the candidate with the replacement and third guard.
- `plan.json`, `capture-freeze.json`, `capture.py`, `build-*.json`, source
  manifests and per-child receipts bind the frozen 16-native/16-allocator
  ABBA experiment. CPU 12; native 100 samples/10 warmups, allocator 3/0.
- `analysis.json` recomputes every raw elapsed statistic and all metric gates;
  `negative-checks.json` rejects four in-memory corruptions and verifies exact
  positive replay. Raw reports, stdout/stderr and before/after identities remain.
- `oracle/` is the unchanged 0712 11-case active-offset probe, deliberately
  retaining its package name and schema. Both compiled reports are byte-identical
  to each other and to 0712. `oracle-freeze.json` binds its three input files.
- `baseline-checks.json`, `candidate-checks.json`, and `quality.json` bind
  coordinator tests to exact source. `quality-initial.json` and
  `quality-second-attempt.json` retain the failed all-target Clippy attempts;
  `quality-exceptions.json` and baseline reproductions identify seven preexisting
  PPTX/XLSB test-only lint errors. Final checks retain warning-denied Clippy for
  all seven libraries and all targets of the other five crates. `review-preflight.json` discloses a
  reviewer’s unplanned pre-capture default-target Cargo execution; it is not
  substituted for those checks. The test author also ran a temporary standalone
  scalar test harness before timing; that temporary binary was removed.
- `preflight-evidence/` and `source-preflight.json` are baseline checks.
  `evidence/` and `source-final.json` are final-source checks. Each has six gates.
- `mechanism-plan.json`, `mechanism-freeze.json`, `mechanism.py` and
  `analyze_mechanism.py` describe the conditional eight Callgrind and sixteen
  RSS children, admitted only after the explicit pilot pass. The raw profiles
  retain all setup, measured and zero-Ir terminal parts. Process RSS peaks are
  individual observations, never sums.
- `mechanism-incomplete-attempt/` retains the first unreceipted profile and
  original scripts after a driver bookkeeping failure; `mechanism-recovery.json`
  records the repair before the full admitted matrix. Native measurements were
  not rerun. `search-attribution.json` is a post-capture exact-edge diagnostic.
- `cleanup.json` retains the six binary identities after removal of the three
  owned scratch roots. Shared preexisting workspace `target/` is not owned
  by this batch and is not removed.
- `audit.py`, `final-report-gate.json` and `artifact-manifest.json` bind terminal
  decisions, exact analyzer replay, final documentation and every packet file.

From the repository root, audit the retained final checkout with:

```sh
python3 -B docs/performance/results/change-0713/audit.py
python3 -B docs/performance/results/change-0713/artifact-seal.py --check
```

The capture drivers intentionally refuse existing outputs. A new experiment
needs a new packet/scratch namespace and source-bound baseline/candidate builds;
these retained results must not be overwritten or cherry-picked. Commands,
source maps, fixture hash, compiler/build logs and host metadata are retained
for reproduction. Callgrind Ir is guest-instruction attribution, not hardware
instructions or native timing. The report makes no cold-cache, throughput,
scaling, cross-architecture or other-format speedup claim.
