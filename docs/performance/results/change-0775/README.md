# Change 0775 evidence

See [the integration report](../../0775-mce-stream-integration.md).
Capture and acceptance status are stated there; an incomplete directory does
not establish a passing gate or measurement claim.

- `integration.json`: original commits, clean cherry-picks and source hashes.
- `review.md`: independent source review and remaining URI-dependent paths.
- `quality.py`, `quality-0/`, `quality.json`: serial owner tests, checks, lint,
  docs, facade/harness build checks and crate boundaries, with source hashes.
- `workspace-Cargo.lock`, `host.json`: workspace resolution and host details.
- `probe-src/`: identical standalone source for both production legs, extended
  from the integrated attribute-bounds harness with review controls.
- `measure-initial.py`, `measure-0/`, `measure-second.py`, `measure-1/`: retained failed oracle attempts.
- `measure.py`, `measure-2/`: final release build custody and serial process captures.
- `probe-corrections.md`: both probe corrections and the existing tree-parser gap.
- `preid.py`, `preid-0/`, `alias-triage.json`: separate text-based reference and exact comparison.
- `comment-followup.json`, `final-source.json`: final source and comment-only follow-up gates.
- `analyze.py`, `analysis.json`: raw replay, input/outcome identity, process-p50
  and peak-RSS statistics, all differential changes retained for triage.
- `validate.py`: offline gate, custody, statistics and optional seal checks.

Native measurements pin each process to CPU 12 and alternate three processes
per leg. Long URI cases use three measured samples after one warmup; ordinary
controls use nine after two. Input generation is outside the timer; parser and
observer work is inside it. RSS includes the whole process. Sample maxima are
not reliable p95/p99 estimates. No physical cold-cache condition is asserted.

Differential results compare the same XML fixtures and generated inputs across
source revisions. Equal failures on generated malformed inputs prove neither
validity nor standards compliance. Alias-pair probes provide separate semantic
checks; any changed behavior must have an explicit disposition.

Reproduction requires isolated worktrees at the archived source references,
the archived workspace locks and standalone probe lock, and updated local
root/target paths in the capture helper. Keep workload execution serial. Never
overwrite a capture folder. Offline replay uses:

```sh
python3 -B docs/performance/results/change-0775/validate.py
```
