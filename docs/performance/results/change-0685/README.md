# Change 0685 evidence packet

Final correctness and evidence audits pass. See the [record](../../0685-xls-worksheet-chain-checkpoints.md)
for measured gains, regressions and the scoped disposition.
Baseline `c207f2c43` contains the 0684 bounded occurrence cache. The candidate
adds one immutable weakly-bound worksheet-start CFB checkpoint per admitted
worksheet index. It retains no source bytes and does not change I/O fences.

The native/counting/allocation/repeat probes are reused unchanged from
`../change-0684/`; their individual Cargo.lock files remain there. This packet's
manifests bind both CFB and XLS Rust sources and repository-relative probe
paths. Reproduce by building the same probe manifests from separate baseline
and candidate worktrees/target directories with release/offline/locked flags,
then passing those binaries to `run-measurements.py` and `run-diagnostics.py`.
Drivers retain exact commands and use CPU 12; adapt CPU and recorded worktree
paths consistently on another host. File source timings use warm OS caches.

`cases.json` freezes twelve cases, including late stored targets in the large
and medium sheets. Baseline A/A and candidate A/B/B/A use 30 fresh-owner samples
per route after three warmups. Setup/source construction precedes open timing;
semantic projection is outside query timers. Never treat diagnostic
`total_elapsed_ns` as API latency. p99 is the maximum of 30 samples, not a
population-tail estimate. Shared-host control drift remains visible.

Allocation probes retain owner/results through capture and prepare the same
coordinate; native prepared routes use a different second coordinate.
Counters and allocation results are separate from native timing. Repeat
process diagnostics use 10/1,010 queries, three repeats, hardware counters and
peak RSS; small or negative subtracted counts are not percentage claims.

All retained bindings, semantic parity, sample structure, repeat allocation
agreement, final quality checks and immutable baseline source are checked by
`audit.py`; `audit-diagnostics.py` checks hardware/RSS evidence. Broad program
work remains active, iWork excluded, `performance_claim: none`.

Final artifacts:

- `comparison.json`: all native phase distributions and 88 allocation groups.
- `timing-diagnostics-comparison.json`: 4,800 additional records, 200 per leg,
  confirming both the large-sheet warm gain and tiny-workflow regression.
- `regressions.md`: every native paired >5% median trigger, with absolute ns.
- `diagnostics-comparison.json`: separate repeat-loop counters and process RSS.
- `profiles-manifest.json`: before/after profile commands and text hashes bound
  through the diagnostic manifests to probes, source, fixture and repeat binary.
- `final-verified/results.json`, `final-builds.json`, `evidence/results.json`:
  final quality/build bindings and dependency/evidence gates.
- `initial-candidate/`: complete initial evidence and five changed source files;
  final restoration avoids the initial temporary weak-reference increment/decrement.

Run `python3 docs/performance/results/change-0685/audit.py` and
`python3 docs/performance/results/change-0685/audit-diagnostics.py` from the
repository root. They also verify binaries when the owned build directories
exist; those large directories and raw perf.data are removed after validation.
Retained profile symbol reports cover user-space attribution; restricted kernel
symbols limit kernel attribution. Profile commands use equal N=2,000,000.
Performance samples use the final source hashes, not the preliminary build logs.

The tiny missing open-plus-three median regresses 5.12–5.65%, tiny stored
4.56–4.74%, and refusal-query phases 5.41–10.55%. These costs are explicit.
A +5.92% short Plan1 file-process RSS flag is absent at the longer repeat length;
no general RSS reduction is claimed. Checkpoint build retention adds 32 measured
bytes, versus 64 conservative logical accounting bytes. I/O/fences are identical.
