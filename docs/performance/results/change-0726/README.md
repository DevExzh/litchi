# 0726 XLS empty-slot setup evidence

**Rejected.** Four native mean checks fail across 4/24 groups. Every native p50
check, all 16 repeated groups, both missing-query benefit gates, allocation,
semantic, source-I/O, budget-fence and binding checks pass. Production is restored
to `3ad29e42da`. `performance_claim: none`; the broader non-iWork goal stays active.

[Performance record](../../0726-xls-empty-slot-setup-pilot.md) explains the
candidate and limitations. All samples, including every failed mean and tail,
are retained.

| Artifact | Purpose |
| --- | --- |
| `hypothesis.md`, `plan.json`, `captures/freeze.json` | Prospective hypothesis, thresholds and frozen identities |
| `candidate-source/`, `candidate-source.json`, `candidate-production.patch` | Exact one-file candidate |
| `*-builds.json`, build logs, `environment.json` | Source, release executable and host bindings |
| `captures/` | Complete A/A + ABBA native, repeat, allocator and budget observations |
| `analysis.json`, `analysis.md`, `analysis-initial.log` | Primary rejected gate result, expected exit 1 |
| `measurement-details.json`, `measurements.md` | All central/tail/drift comparisons and four failed mean checks |
| `audit.py`, `audit-review.md`, `audit-initial.log` | Independent source, custody and statistical audit |
| `distribution-review.md` | Independent failed-mean/tail attribution and prospective follow-up boundary |
| `source-review.md` | Actual source diff and existing test coverage review |
| `*-profile-*`, `*-profile.tsv`, `*-profile.stderr` | Separate two-million-query missing-result perf captures |
| `corpus/`, `corpus-comparison.json` | Exact real and generated corpus comparison |
| `integration/`, `qualification.md`, `test-counts.json` | Final correctness qualification |
| `evidence/` | Six repository boundary, claim and coverage checks |
| `negative-checks.py`, `negative-checks.json` | Seventeen controls bound to the frozen analyzer |
| `cleanup.json`, `terminal-replay.json`, `*-final.log` | Exact removed executable witnesses and offline replay |
| `disposition.json`, `final-state-check.log` | Final baseline source, constraint and cleanup verification |
| `artifact-manifest.json` | Terminal path, byte-count and SHA-256 census |

Native capture includes 24 groups × six legs × 100 fresh measured owners × eight
queries, plus three warmup owners per group/leg. Repeat capture includes 16
groups × six legs × nine processes × 50,000 queries. Allocator capture includes
96 groups × two phases × three repeats. Instrumented timing is never pooled
with native results. There were no failed/replaced main captures.

For fresh reproduction, use an isolated checkout of `3ad29e42da`, matching
fixtures and unchanged 0684/0686 probes. Build baseline, then install the one
archived candidate file, build candidate, qualify, freeze and capture in a new
output directory. Drivers refuse overwrites. Do not substitute a new run for
this packet's observations.

For offline replay, use a copy of the packet and install the archived candidate
in an isolated checkout. The frozen analyzers require its source census even
though production was restored. Exact cleanup witnesses replace removed binary
files. Run:

```text
python3 -B docs/performance/results/change-0726/analyze.py --offline
python3 -B docs/performance/results/change-0726/audit.py --allow-rejected
python3 -B docs/performance/results/change-0726/audit-corpus.py
python3 -B docs/performance/results/change-0726/negative-checks.py
```

The first two commands must exit **1** for verified rejection; the others exit
zero. Thresholds stay frozen. Analysis timestamps may change during replay;
`terminal-replay.py` preserves recorded primary-result bytes and compares gate
results. Restore baseline afterward. The current restored checkout can directly
run `final-state-check.py` and `artifact-seal.py --check` without reinstalling
the candidate. Never overwrite a sealed packet to test reproduction.
