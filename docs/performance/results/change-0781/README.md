# Change 0781 evidence packet

This packet records the fresh PPT text-ownership experiment and the repair of
the default performance CI matrix. The `Cow<str>` production candidate is
rejected; the live PPT source is restored to base commit
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8`. The CI repair is retained. No
retained production speedup, CRUD-row promotion, or full-goal completion follows.

The capture-time roots were:

```text
worktree: /home/zhuhe/code/litchi-worktrees/0781-ppt-borrowed-text
packet:   docs/performance/results/change-0781
target:   /home/zhuhe/code/litchi-target-0781
```

The active scope is OLE2/OOXML. ODF optimization is deferred under the 0758
owner decision, and iWork is excluded.

## Retained evidence

| Area | Evidence | Cardinality or result |
| --- | --- | ---: |
| Frozen design | [`design.md`](design.md), [`plan.json`](plan.json), [`architecture-inputs.json`](architecture-inputs.json) | 10 cases; 35 architecture inputs unchanged |
| Probe and oracle | [`probe-src/`](probe-src/), [`probe-tests/`](probe-tests/), [`qualification-review.md`](qualification-review.md) | 31 default-feature + 10 all-feature tests pass |
| Initial correction | [`initial-probe/`](initial-probe/), [`qualification-initial/`](qualification-initial/), [`qualification-correction.json`](qualification-correction.json) | 64 authored trailing spaces retained; failed setup retained |
| Qualification | [`qualification/`](qualification/), [`qualification-parity.json`](qualification-parity.json) | 10 before-only reports; six initial/corrected output identities unchanged |
| Native matrix | [`native/`](native/), [`native-summary.csv`](native-summary.csv), [`native-pairs.csv`](native-pairs.csv) | 120 processes; 3,600 samples; 20 before/after process groups |
| Allocation matrix | [`allocation/`](allocation/), [`allocation-summary.csv`](allocation-summary.csv) | 40 processes; 120 samples; 13 allocator fields |
| Observer lanes | [`perf/`](perf/), [`heaptrack-before/`](heaptrack-before/), [`heaptrack-after/`](heaptrack-after/), [`observer-analysis.json`](observer-analysis.json) | 12 whole-process perf receipts; one three-sample Heaptrack process per leg |
| Candidate custody | [`candidate/`](candidate/), [`disposition.json`](disposition.json), [`restored-source.json`](restored-source.json) | 12 candidate paths archived; production change rejected |
| PPT quality | [`quality.json`](quality.json), [`test-summary.json`](test-summary.json) | six gates; 1,286 passed, 0 failed, 11 ignored, 34 suites |
| CI repair | [`ci-capture/`](ci-capture/), [`ci-quality-1/`](ci-quality-1/), [`ci-analysis.json`](ci-analysis.json), [`ci-audit/`](ci-audit/) | smoke 41 rows; full 213 rows; 75 tests; six gates |
| Analysis and flags | [`analysis.json`](analysis.json), [`all-flags.csv`](all-flags.csv), [`results-review.md`](results-review.md), [`results-audit.md`](results-audit.md) | 170 reports; 3,730 samples; 20 paired and 42 spread native flags; zero allocation flags |

The strict raw oracle checks untrimmed OfficeArt `ClientTextbox` atoms with
bounds, UTF-16/ASCII validity, record traversal, authored trailing spaces,
and rich separators. It is separate from the public reader's documented
per-atom trimming oracle and is not an independent CFB parser. The final raw
checks pass for all 170 primary reports and 3,730 samples.

The decisive native p50 results are `many/write +12.121%` and
`many/lifecycle +10.551%`, both across all six paired blocks. Payload improves
by 25.555% and 19.186%; Unicode improves by 8.314% and 8.073%. Allocation
requested bytes fall on those large fixtures, while peak-above-entry and net
live remain unchanged. Heaptrack attributes 192 conversion events and
7,680,000 requested bytes before the candidate, versus zero after; this is
allocation-site evidence and carries no timing, RSS, physical-copy, or causal
claim.

The CI repair replaces stale 37-case/201-row workflow assertions with the
manifest-backed 41-case/213-row contract, and uses active CRUD index v2. The
full capture binds 43 corpora and 213 report-to-corpus bindings. It does not promote CRUD
coverage or expand ODF/iWork scope. The CI review records a shared provenance
follow-up for the stale `main.rs:Case::DEFAULT` literal; the declaration is in
`lib.rs`.

## Offline replay and reproduction

From the repository root:

```bash
python3 -B docs/performance/results/change-0781/validate.py
python3 -B docs/performance/results/change-0781/tables.py --check
```

The final validator requires `seal.json` and checks every packet file, the
before/candidate/restored source archives, exact commands and raw reports,
probe tests, CI replay, initial binary relocations, observer analyses, and
cleanup witnesses. It works after removal of the original worktree and target.

For a fresh measurement, follow the ordered reproduction instructions in the
[0781 report](../../0781-ppt-text-ownership-and-ci-matrix.md). Use new output
folders and a new target, the same base, plans, probe and locks, and the archived
applied patch. The scripts refuse existing capture/build folders. Generate the
new analyses and tables before sealing that new packet; never rewrite retained
receipts to make the new run appear to be this one.

Seven executable identities were verified before the owned target was removed
(3,248,857,328 file bytes); post-removal replay passes. The final commit and
owned-worktree/branch cleanup are recorded in the report outside this sealed
packet.
