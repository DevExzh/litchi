# 0522: rejected common-cell scanner candidate; retained guards

The production candidate was rejected under the frozen native admission rule.
Dense-sparse primary p50 improved in both repeats, but medium moved +1.29%
and -1.30%. Instruction and allocation reductions did not establish useful
repeatable primary total improvement. The final checkout restores the baseline
scanner and retains the new codec tests, opt-in noncompact guard and evidence.
See [the decision record](../../changes/0522-xlsx-cell-reference-guard.md).

This campaign measures a source-local XLSX scanner change against revision
`96fe51a3a70d97dfc49610788bf32963ae58ea2e`. The address scan retains its checked
attribute traversal, decoding and coordinate errors, and also proves whether
the existing compact tag representation is sufficient. Only that proven case
omits the second tag scan. Prefixed and additional-attribute tags retain the
existing full tag path. Candidate XML validation and semantic readback remain.

## Frozen inputs and execution

`plan.json` and `run.py` freeze the hypothesis, setup, guards, admission rule
and capture protocol before the first build/capture receipt. Both stages use
identical tests and the new opt-in `noncompact` harness shape. This shape has
the medium geometry and values, with alternating unprefixed numeric cells
carrying `r` plus `t="n"`, and prefixed `x:c` cells carrying `r`. Every sheet
contains both forms; the namespace is bound at the worksheet root.

Each stage retains 820 native samples: four primary children with 20 warmups
and 100 samples, plus fourteen guard children with 10 warmups and 30 samples.
The native block order is A1/B1/B2/A2. A1 is captured before applying the
candidate; final A2 executes the unchanged retained baseline executable after
the candidate runs. Its receipt separately binds the baseline compiled-source
manifest and the current candidate working-source manifest. The capture driver
requires the complete manifests to differ only in the scanner source file,
and checks executable hashes before and after every child.

Separate allocator stages retain 20 samples each in four children, without
warmup. Separate Callgrind stages retain four one-sample children each. All
builds, captures and quality gates run in one serial lane. This controls campaign processes and does not assert
exclusive use of the shared host. Candidate quality
receipts bind the measured candidate. A final focused codec check binds the
restored baseline source after rejection. Builds use two
jobs, no incremental compilation, the standalone harness workspace's lockfile
and ordinary release defaults. Root-workspace LTO settings do not apply.

For reproduction, use a new evidence directory and frozen plan; the driver
refuses existing receipts. Start from the recorded revision and apply the
retained `baseline/source.patch`, which supplies the tests and noncompact
guard. Retain an external copy of this bundle when checking out that revision.
Use the following sequence, applying `cell-reference-candidate.patch` only
after all preceding baseline commands complete:

```sh
python3 -B docs/performance/results/change-0522/run.py baseline freeze
python3 -B docs/performance/results/change-0522/run.py baseline build-normal
python3 -B docs/performance/results/change-0522/run.py baseline native-r1
python3 -B docs/performance/results/change-0522/run.py baseline profile
python3 -B docs/performance/results/change-0522/run.py baseline build-alloc
python3 -B docs/performance/results/change-0522/run.py baseline alloc
# Apply the production candidate, then:
python3 -B docs/performance/results/change-0522/run.py candidate freeze
python3 -B docs/performance/results/change-0522/run.py candidate build-normal
python3 -B docs/performance/results/change-0522/run.py candidate native-r1
python3 -B docs/performance/results/change-0522/run.py candidate native-r2
python3 -B docs/performance/results/change-0522/run.py baseline native-r2
python3 -B docs/performance/results/change-0522/run.py candidate profile
python3 -B docs/performance/results/change-0522/run.py candidate build-alloc
python3 -B docs/performance/results/change-0522/run.py candidate alloc
```

After capturing and checking the candidate, reverse only the production
patch to restore the baseline scanner. `disposition.json` binds that final
source state to the rejection, and `baseline/check-final-codec.*` retains the
19-test final check. `final-source-review.md` reviews the measured candidate
before rejection; it does not describe an adopted production change.
The verifier checks both captured source patches against the recorded Git
revision and checks the final checkout against the baseline manifest.

The numerical wrapper reuses the retained 0521 verifier; the profile wrapper
reuses 0521 custody logic and 0519 raw parsers. Helper hashes are retained in
the reports. Recompute without the removed build binaries:

```sh
python3 -B docs/performance/results/change-0522/analyze.py /tmp/0522-comparison.json
python3 -B docs/performance/results/change-0522/analyze_profiles.py /tmp/0522-profiles.json
python3 -B docs/performance/results/change-0522/verify.py --sealed
```

Profile annotations regenerate with fixed Perl hash ordering. Only the fourth
numbered dump measures the selected timed commit; the earlier lifecycle dumps
and final zero-Ir dump remain retained. Inner call metadata can include work
outside collection and cannot establish event or allocation counts.

## Boundaries and correctness

Native total time sums open, planning, staged sets/commit, and publication.
Publication includes dropping its returned snapshot. Sink setup, other handle
destruction, reopen and oracles are outside that sum. Acquisition-order phase
vectors reconcile exactly through the sorted total's `sample_order`.

The canonical allocator region covers staged sets and commit only; its build
timings never enter native comparisons. Incremental peak subtracts entry live
bytes from the region peak. Whole-child RSS includes setup and verification.
Source input is instrumented in-memory `ReadAt` with a fresh editor/cache each
iteration; printed filesystem/range defaults do not activate those providers.

The baseline-compatible codec tests cover compact and fallback tags together,
exact typed/display error precedence, and genuine truncation. The initial
truncation fixture accidentally appended closing markup and was corrected
before freezing. The new harness test initially required an explicit `usize`
counter annotation; that compile failure and corrected pass are also retained.
Neither issue changed production code or frozen captures.

Native oracles cover deterministic output, untouched-member preservation,
no-op, clear/remove, foreign/stale refusal and semantic inverse restoration.
The vendor-extension shape additionally exercises partial-sink failure.
The primary one-percent shapes touch all four sheets, so zero unselected-sheet
reads does not demonstrate selective-subset access. The vendor fixture tests
unknown package members, not unknown worksheet grammar.

No native Office-producer, fuzz, physical-provider, cold/range, hardware-counter,
parallel-scaling or broad coverage-catalog completion is asserted. OLE2/OOXML
remains the active priority; ODF is deferred until its optimization goal is
complete, and iWork is excluded.
