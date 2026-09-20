# Change 0704 — bounded PPTX retained projection

This packet measures a PPTX-only candidate that retains immutable processed
slide projections across a changed opened-presentation commit.  It is a
bounded experiment: the planned admission budget is 1 MiB, and the candidate
must expose an explicit fallback when the aggregate retained projection
budget is unavailable or exceeded.  The packet contains no shared MCE oracle,
declaration controls, assembly capture, or copied measurement output from an
older batch.

The baseline is the fresh 602-file source census at the revision recorded in
`baseline.json`.  The candidate may add or modify Rust under
`crates/litchi-pptx`, including a memo module and focused tests.  The source
census and both build drivers reject changes under `litchi-ooxml-common` or
`litchi-opc`; the audit does not require any shared codec change.  Cargo
manifests, lockfiles, constraints, and probe sources are frozen for both
phases.

The native probe package is named `probe0704` and keeps the established
workflow names `one`, `noop`, and `two`.  Its edit marker remains
`litchi-perf-0691-opened-phase`; semantic reopen, revision, untouched-part,
and target-text checks are performed outside the timers.  The refusal probe
package is `probe0704-refusal`, while its output identity and MCE fixture
marker remain `0693-refusal` so the old error-precedence coverage stays
comparable.

The final candidate passes 26 focused tests and all seven integration gates.
Initial and longer ABBA results, exact work-removal counts and memory costs
are reported in [the change record](../../0704-pptx-bounded-slide-mce-retention.md).
The follow-up uses all 13 native workflows and ten refusal cases at 300
samples/ten warmups; see [followup-README.md](followup-README.md). Separate
[mechanism traces](mechanism-README.md) prove real changed commit removes
twelve/eleven default MCE calls without altering native timing binaries.

## Reproduce

Run these commands from a disposable checkout with this packet at its normal
relative path.  Preserve the root `Cargo.lock`; the two standalone probe
lockfiles are already frozen and validate with `cargo metadata --locked`.

```sh
python3 docs/performance/results/change-0704/prepare-control.py
python3 docs/performance/results/change-0704/check-control.py
python3 docs/performance/results/change-0704/source-census.py baseline
python3 docs/performance/results/change-0704/build.py baseline
python3 docs/performance/results/change-0704/build-refusal.py baseline
python3 docs/performance/results/change-0704/measure.py baseline
python3 docs/performance/results/change-0704/measure-refusal.py baseline
python3 docs/performance/results/change-0704/measure-allocations.py baseline
python3 docs/performance/results/change-0704/profile.py baseline
```

The native matrix has 13 workflows over the real LibreOffice deck, the
same-length marker mechanism control, generated `12x8`, and two notes-bearing
decks.  It uses two baseline A/A legs and two candidate ABBA legs, reversing
case order on odd legs, with 100 samples and five warmups per process.  The
refusal matrix has ten bounded cases with the same schedule.  Allocation
runs are diagnostic and are never folded into native timing claims.

After the candidate source is present, run the source boundary and its
dedicated threshold tests before freezing candidate binaries:

```sh
python3 docs/performance/results/change-0704/source-census.py candidate
python3 docs/performance/results/change-0704/run-functional-tests.py
python3 docs/performance/results/change-0704/build-retention.py
python3 docs/performance/results/change-0704/run-retention.py
python3 docs/performance/results/change-0704/build.py candidate
python3 docs/performance/results/change-0704/build-refusal.py candidate
python3 docs/performance/results/change-0704/measure.py compare
python3 docs/performance/results/change-0704/measure-refusal.py compare
python3 docs/performance/results/change-0704/measure-allocations.py candidate
python3 docs/performance/results/change-0704/profile.py candidate
```

`functional-test-inventory.json` names the new retention module and its
explicit focused test inventory.  `run-functional-tests.py` uses the
`mce_retention` filter, which also captures the two focused policy-intersection
tests, and records every passing test name. The old generic package suite is
not sufficient evidence for the 1 MiB threshold or pressure fallback.

The retention observer is a candidate-only public-API lane.  It records raw
retained-byte observations at capture, clone, transaction, commit, and
publication for the real marker-bearing deck and marker-free `generated:12x8`.
Its standalone lockfile, production source maps, binary hash, and raw outputs
are bound by `retention-build.json` and `retention-runs.json`.  It makes no
timing, allocator, RSS, or throughput claim.

The optional `mechanism/` lane is a separate source-only diagnostic.  It may
temporarily instrument the shared MCE codec only after the native binaries and
measurements are frozen.  Its build receipt must bind the instrumented patch,
the pre-instrumentation candidate source map, the restored and post-restore
shared-source hashes, and its dedicated binary.  The normal candidate census
continues to require unchanged shared OOXML and OPC sources; the mechanism
exception applies only inside its temporary build/trace window and makes no
timing claim.  Its binary is a separately named entry under
`../litchi-0704-bin`; cleanup binds and removes that binary with the shared
0704 target/bin scratch only after the restored-source checks pass.

When this lane is enabled, run it only after the candidate native matrix is
terminal:

```sh
python3 docs/performance/results/change-0704/mechanism/build.py
python3 docs/performance/results/change-0704/mechanism/run.py
```

The run produces twelve fresh process traces (real/generated × noop/one/two ×
two repeats).  `analyze_trace_0704.py` records direct default-MCE call counts
from the raw stderr traces; those counts are diagnostic evidence and carry no
timing or allocator interpretation.

The profile uses the `commit` prefix, not the capture-only prefix.  Its
denominator includes open, presentation capture, working clone, one text
edit, and commit recapture, and stops before publication and serialization.
The native phase timers remain the five named calls capture, clone, set-text,
commit, and apply.  Keep these denominators separate.

Derive descriptive summaries only after both phases have terminal raw output:

```sh
python3 docs/performance/results/change-0704/summarize.py
python3 docs/performance/results/change-0704/summarize-allocations.py
python3 docs/performance/results/change-0704/summarize-refusal.py
python3 docs/performance/results/change-0704/report-metrics.py
python3 docs/performance/results/change-0704/run-integration.py
python3 docs/performance/results/change-0704/quality-summary.py
python3 docs/performance/results/change-0704/run-evidence.py
python3 docs/performance/results/change-0704/scripts-manifest.py
```

The seven integration gates and six evidence gates run after the measurement
lane.  Run the terminal audit before cleanup, then remove only this batch's
owned scratch and seal the retained receipts:

```sh
python3 docs/performance/results/change-0704/audit-statistics.py
python3 docs/performance/results/change-0704/audit_followup.py
python3 docs/performance/results/change-0704/audit.py
python3 docs/performance/results/change-0704/cleanup.py --apply
python3 docs/performance/results/change-0704/seal.py
python3 docs/performance/results/change-0704/artifact-seal.py --write
python3 docs/performance/results/change-0704/artifact-seal.py --check
```

The final source witness in `candidate-source-witness.json` binds
`final_candidate_patch`, its SHA-256, and the complete
`final_candidate_source_sha256` map to a full `d48523eec2` to candidate diff
covering every changed or added PPTX production and test source.  The isolated
coder handoff is retained under `preflight/` as historical evidence; it is not
the final candidate source.  The owned scratch paths are
`../litchi-target-0704`,
`../litchi-0704-bin`, `../litchi-0704-profile`,
`../litchi-0704-mce-candidate`, and this packet's generated
`marker-control.pptx`.  The candidate worktree is removed only after
the historical handoff patch and witness have been verified; its source
hashes are retained in `cleanup.json`.  The root
`Cargo.lock`, standalone probe lockfiles, raw packet evidence, and unrelated
scratch remain.

## Interpretation

Allocation records retain raw allocation calls, requested bytes,
reallocation bytes, phase baseline/current live bytes, and phase peaks.  The
candidate is expected to differ: retained projection memory is the subject of
the experiment.  Report raw values and the observable retained/current-live
and peak deltas separately; they are allocator observations, not an RSS
bound.  Cache metadata, allocator rounding, transient overlap, and any
unobservable internal accounting are not inferred from these counters.

The successful native probe excludes source loading, target discovery,
serialization, reopen, and semantic oracle work from its timers.  The
refusal probe separately checks ten error and valid controls for typed error
identity and stable graph fixtures.  The marker-control archive changes the
namespace URI and is a mechanism control, not semantic equivalence or a
preservation oracle.  No cold-cache, native Office, concurrent, cross-platform,
or general save claim follows from this packet.  iWork is outside its scope.
