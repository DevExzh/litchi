# 0725 revised XLS checkpoint evidence

**Rejected.** Native timing fails 17/24 groups; repeated timing passes 16/16.
Both late-target benefit gates and all allocation, semantic, counted-I/O,
budget-fence and binding gates pass. Production is restored to `ee5e0b0650`.
The broader non-iWork goal remains active. `performance_claim: none`.

[Performance record](../../0725-xls-revised-checkpoint-pilot.md) explains the
candidate, decisive regressions and limitations. All observations are retained.

| Artifact | Purpose |
| --- | --- |
| `hypothesis.md`, `plan.json`, `captures/freeze.json` | Prospective hypothesis, thresholds and immutable capture identities |
| `candidate-source/`, `candidate-source.json`, `candidate-production.patch` | Exact rejected source/test bytes |
| `*-builds.json`, build logs, `environment.json` | Release executable, source and host identities |
| `captures/` | Native A/A + ABBA, repeated-loop, allocator and budget raw captures |
| `analysis.json`, `analysis.md`, `analysis-initial.log` | Frozen primary gate result; expected analyzer exit 1 |
| `measurement-details.json`, `measurements.md` | Every central/tail/drift comparison, including 102 failed central checks |
| `audit.py`, `audit-review.md`, `audit-initial.log` | Independent source, custody and statistic verification; rejection remains exit 1 |
| `trace/`, `trace-audit.log` | Separate instrumented route and ordered-query evidence |
| `*-profile-*`, `*-profile.tsv`, `*-profile.stderr` | Separate 2-million-query sampled diagnostic profiles |
| `corpus/`, `corpus-comparison.json` | Exact 126-fixture owned/file and 70,001-cell generated comparisons |
| `integration/`, `qualification.md`, `test-counts.json` | Final candidate correctness checks |
| `evidence/` | Six repository boundary, claim and coverage gates |
| `negative-checks.py`, `negative-checks.json` | Fourteen controls bound to the frozen analyzer |
| `initial-candidate-build-failure/` | Archived unused-mut build failure before correction |
| `pre-consolidation-qualification/` | Passing prior test inventory, before duplicate-test removal |
| `cleanup.json`, `terminal-replay.json`, `*-final.log` | Exact deleted executable witnesses and offline replay results |
| `restoration.json`, `restoration-check.log` | Exact baseline production restoration |
| `artifact-manifest.json` | Terminal path, byte-count and SHA-256 census |

Native capture includes 24 groups × six legs × 100 fresh measured owners × eight
queries, plus three warmup owners per group/leg. Repeated capture includes 16
groups × six legs × nine processes × 50,000 measured queries. Allocator capture
includes 96 groups × two phases × three repeats. Diagnostic traces and profiles
do not contribute timing samples. No failed main capture was replaced.

For fresh reproduction, use an isolated checkout of `ee5e0b0650`, this packet,
matching fixtures and unchanged 0684/0686 probes. Build baseline before installing
the three archived candidate files; build candidate, qualify, then freeze and
capture in a fresh output directory. Drivers refuse capture overwrites. This
packet's measurements must not be replaced by another run or machine.

For offline analysis, install the archived candidate in an isolated checkout;
the frozen analyzer requires its source census even though production was
restored. Exact `cleanup.json` witnesses allow removed binaries. Run:

```text
python3 -B docs/performance/results/change-0725/analyze.py --offline
python3 -B docs/performance/results/change-0725/audit.py --allow-rejected
python3 -B docs/performance/results/change-0725/trace-analyze.py
python3 -B docs/performance/results/change-0725/audit-corpus.py
python3 -B docs/performance/results/change-0725/negative-checks.py
```

The first two commands must exit **1** for the verified rejection; the remaining
commands exit zero. Thresholds remain frozen. Generated analysis timestamps can
change during replay: use a copy instead of overwriting sealed evidence.
`terminal-replay.py` preserves the recorded primary result bytes and compares
independent hard-gate results before/after cleanup. Restore baseline after replay.

The final restored checkout directly supports `restoration-check.py` and
`artifact-seal.py --check`. These verify exact baseline sources, unchanged
constraints, archived candidate identity, rejected disposition, cleaned roots,
and packet bytes without reinstalling the candidate.
