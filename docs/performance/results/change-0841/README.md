# Change 0841 evidence packet

Current-source PPT instruction diagnosis of the rejected 0782 borrowed-text
candidate. Production is restored exactly to base `307434b454`. There is no
adopted optimization or coverage promotion. All 156 workload reports / 3,636
samples validate; none was discarded, replaced or rerun.

The historical many/Unicode regressions do not reproduce in these binaries.
Guest Ir falls in all four profiled cases, while the Unicode textbox encoder's
self Ir is unchanged. Tiny native write increases 2.46%; all ten cases and all
seventy native metric rows remain visible. No native p50/RSS threshold or
block-p50 spread flag fires. These findings do not establish a cause for 0782
or authorize adopting this representation without a fresh ordinary-route trial.

From the recorded checkout, replay without rebuilding:

```bash
python3 -B docs/performance/results/change-0841/reader_tests_final.py
python3 -B docs/performance/results/change-0841/analyze_v3.py
python3 -B docs/performance/results/change-0841/crosscheck.py
python3 -B docs/performance/results/change-0841/interval_crosscheck.py
python3 -B docs/performance/results/change-0841/tables.py
python3 -B docs/performance/results/change-0841/close.py verify
```

`close.py verify` additionally checks current production, normative and unrelated
file identities and the exact commit parent/path set. The packet preserves the
recorded absolute checkout paths; it is not a path-independent replay bundle.
The pure readers/reducers require the retained historical 0782 fixture seal,
0840 CFB parser and 0784 Callgrind parser source too. Raw data and every source
input are identified, rather than substituting current outputs for old reports.

## Fresh measurement protocol

Use a fresh checkout of the recorded base, separate empty evidence/build/scratch
directories, fresh owner markers, and new origin/source/host receipts. Do not run
the historical mutation/build/capture drivers over this sealed packet. Their
exclusive creation and identity guards intentionally refuse reuse. Adapt only
the paths and fresh custody records, reviewing the resulting plan before capture.

1. Preserve the final `probe/` sources, root-aligned `probe/Cargo.lock`, and
   `candidate.patch`. Read `plan.json`, `selection.md`, `review.md`, and the
   final source/reader repair records. Capture the new host, source, normative
   and unrelated-file inventory before mutation; create the marker-owned target
   and scratch roots used by `driver.py`.
2. Run `quality.py before`, `build.py before`, and `symbols.py before` against
   baseline production. The driver fixes the recorded environment, including
   opt3/fat-LTO, one codegen unit, two build jobs and no ambient RUSTFLAGS.
3. Run baseline qualification for all ten cases. Verify exact sealed fixture
   identities and independently parsed CFB bytes before `candidate.py apply`.
4. Apply the exact archived patch, then run candidate owner/probe quality,
   build/symbol qualification and all ten candidate qualification cases. Keep
   equal-length retained binary paths. Compare both legs' output bytes exactly.
5. Before any formal run, qualify the final reader against all artifacts and
   malformed report/profile inputs. Use the corrected `readers_v2.py`,
   `qualify_v3.py`, `analyze_v3.py` semantics from the outset in a new run; do not
   reproduce the historical reader mistakes. Freeze all input hashes and the
   complete schedule before looking at performance.
6. Run the 120 native and sixteen Callgrind processes in `plan.json` serially.
   Exact actual argv and terminal receipts are retained under `commands/` and
   `runs/`. Callgrind uses collect-at-start off, exact owner toggle, zero-before
   and dump-after. Each profile has one measured operation and no warmup.
7. Replay all native values, profile cost conservation, archive identities,
   command/source/binary bindings, raw quantiles and paired intervals. Report
   every case and limitation. Restore production exactly regardless of result:
   this diagnostic's plan does not permit adoption. Hash-verify and remove only
   owned temporary roots, seal, and commit the evidence.

## Evidence and corrections

- `plan.json`, original/additive admissions and freeze/build receipts bind both
  source states, probes, locks, emitted owner code and fixed schedules.
- `runs/` retains every JSON report, time counters, qualification PPT and raw
  Callgrind publication. The JSON/text oracles are separate from the independent
  CFB walker. Every output digest matches the historical golden and the retained
  qualification artifact.
- `native.md` / `native.csv` retain all seventy individual native metric rows.
  `profiles.md` and `analysis.json` retain both instruction repeats, conserved
  function self-cost differences and non-additive direct-edge diagnostics.
- `reader-freeze-violation.json` records an agent's post-freeze reader edit,
  detected before profile launch. The exact original was restored, the observed
  replacement is archived, and corrected `readers_v2.py` is admitted additively.
  All twenty qualification rows and 120 native report reductions are identical.
- Original fifteen-mutation checks and corrected twenty-one checks remain.
  `reader_tests_final.py` also fixes the v2 suite's inherited receipt label to
  bind the actual imported reader. One of the twenty-one checks demonstrates
  the original reader refusing a valid child-only command header; it is not a
  malformed-input case. No test receipt was deleted or overwritten.
- `analysis-repair.json` retains the first reducer's header-spacing failure and
  the additive `.lstrip()` correction before exact child-argv comparison.
  No numerical method, metric, schedule or threshold changed.
- `probe-review.md` records the initial lint/global-allocator test failures and
  dependency lock alignment before any release workload build.
- `compression-review.md` and `compression-feature-evidence.json` retain the
  separate flate2 1.1.9/1.1.10 runtime-dispatch distinction. No dependency changes
  or compression performance result are claimed here.

Final source restoration, binary cleanup and the commit seal are separately
verified. The broader OLE2/OOXML performance goal remains active.
