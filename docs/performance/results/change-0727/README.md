# 0727 warm-tail replication evidence

Diagnostic only. The exact rejected 0726 candidate remains unretained; production
source is unchanged. Three fixed cycles retain all 324 processes, 32,400 measured
owners and 259,200 queries. Original focus metrics produce 14 failed central
checks; all seven metrics across all six cells produce 60. No measurement is
removed or used to retroactively accept 0726. `performance_claim: none`.

[Performance record](../../0727-xls-warm-tail-replication.md) explains recurrence,
controls and limits. The broader non-iWork program remains active.

| Artifact | Purpose |
| --- | --- |
| `hypothesis.md`, `plan.json`, `freeze.json` | Prospective process matrix, unchanged diagnostic thresholds and frozen identities |
| `sources/`, `source-*.json`, `builds.json`, build logs | Exact baseline/candidate ancestry and fresh release binaries |
| `build-restoration.json`, `source-guard.py` | Exact baseline restored before capture and after the experiment |
| `environment.json`, `constraints.json` | Host and unchanged GOAL/ADR bindings |
| `captures/manifest.json`, raw JSON/stderr | Every command, process identity, output and hash |
| `analysis.json`, `analysis.md`, `report-summary.json` | Complete per-process statistics, all comparisons, failed checks and recurrence summaries |
| `audit.py`, `audit-review.md`, `audit.json` | Independent source, command, raw outcome and statistical audit |
| `negative-checks.py`, `negative-checks.json` | Ten controls of actual analyzer behavior |
| `priority-review.md`, `priority-review-initial.md` | Corrected broader queue review and archived stale initial assessment |
| `cleanup.json`, `terminal-checks.json`, terminal logs | Exact executable witnesses and post-cleanup replay |
| `artifact-manifest.json` | Exact terminal packet census |

Each of three cycles uses aa1, aa2, a1, b1, b2, a2. Each of six cells has three
fresh processes per leg, each with three warmup owners and 100 measured owners
taking eight queries. Each process is the replication unit. All seven native
metrics and every p50/mean/p95/p99/maximum comparison are retained in JSON;
no pooling of owner samples establishes a stronger replication count.

Fresh reproduction requires matching fixtures, unchanged 0686 probe sources,
and an isolated baseline checkout of `959daa11e5`. `build.py` temporarily installs
the exact 0726 archived source, builds both variants serially, and restores the
original source in `finally`. Use a new output directory: freeze and capture
refuse overwrites. Do not replace this packet with another run.

Offline replay needs the restored baseline source and fixtures, not the
candidate installed. Removed binaries require exact cleanup path/hash/size
witnesses. Use a copy of the sealed packet when running scripts that rewrite
reports:

```text
python3 -B docs/performance/results/change-0727/source-guard.py
python3 -B docs/performance/results/change-0727/analyze.py
python3 -B docs/performance/results/change-0727/report-tables.py
python3 -B docs/performance/results/change-0727/audit.py --allow-rejected
python3 -B docs/performance/results/change-0727/negative-checks.py
```

The independent audit's expected exit is **1**, preserving its fail-closed timing
flag; all other commands above exit zero. Its custody, semantic and statistic
gates pass independently of `timing=false`. Neither interface has authority to
retain the candidate. `terminal-checks.py` checks expected exit codes and requires
exact report hashes before/after replay, then checks report classification,
structural claims and the non-iWork boundary. The read-only
`artifact-seal.py --check` verifies every packet byte.
