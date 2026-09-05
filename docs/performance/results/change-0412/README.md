# Change 0412 evidence

Observer-only ABBA comparison and a separate plain `OwnedSource` baseline.
No production speedup claim. Control is `70756ae67e6763428759e8f446718ce68a528976`;
measured candidate is `63c95bc22d5883c8ecab0872030757e5584254f7`.

- `capture/protocol.json` declares the scenarios, sampling and measurement scopes.
- Two build identities bind clean source revisions, flags and executable hashes.
- `capture/commands.json` records all 18 sequential capture commands.
- Normal and allocator captures each use control/candidate/candidate/control order.
- Four plain-source normal and two allocator children establish the new baseline.
- One CPU profile and three PMU captures observe the whole plain-source process.
- `capture/0412-comparison.json` checks sample order, semantic oracles and locality.
- `capture/schema-verification.json` applies repository schema/catalog checks.
- `capture/uncertainty.json` retains within-child median intervals and eager review triggers.
- `capture/resources.json` retains 14 RSS observations, profile identity and PMU counts.
- `attribution/` contains period-weighted whole-process and production-ancestor exports.
  The initial unverified report is retained as a diagnostic; the unsuffixed report
  and published replay pass the final required-identity checks.
- `checks/` contains passing gates, initial diagnostics and corruption rejection probes.

Normal and counting-allocator timings are separate observations. Input cloning,
source construction, semantic/locality validation and owner drop are outside the
operation timer/allocation regions. They remain in whole-process RSS, PMU and
CPU profiles. Warmups and samples share each child's heap. No cold-file,
physical-I/O, remote-source, scaling, exact L1/LLC or operation-peak claim follows.

The XLS observer version intentionally differs. The general comparator rejects
that difference; this bundle's observer-only verifier permits precisely the
result-level XLS scope marker while requiring logical evidence equivalence.
Executable hashes were checked against the binaries at capture time. Replay
binds the retained build/journal/report records; the large executable files are
not published. Rebuilding may have a different binary hash and creates new evidence.

From a checkout containing this bundle, with Python and `zstd` installed:

```sh
python3 docs/performance/results/change-0412/verify-artifacts.py
python3 docs/performance/results/change-0412/verify-capture.py --root docs/performance/results/change-0412/capture --output /tmp/litchi-goal-0412-reverified.json
python3 docs/performance/results/change-0412/schema-check.py --root docs/performance/results/change-0412/capture --repo-root . --output /tmp/litchi-goal-0412-schema-reverified.json
python3 docs/performance/results/change-0412/verify-resources.py --root docs/performance/results/change-0412/capture --output /tmp/litchi-goal-0412-resources-reverified.json
python3 docs/performance/results/change-0412/replay-attribution.py --repo-root . --output /tmp/litchi-goal-0412-attribution-reverified.json
```

The artifact manifest binds every published file and both representations of
losslessly compressed originals. Reports and raw perf/log text retain their
original bytes; capture paths remain historical identity fields. Attribution
replay materializes inputs privately and requires Git access to the captured
revision for source binding. No benchmark or perf capture runs during replay.

For new measurements, adapt the retained build/capture scripts' absolute paths,
use clean detached worktrees at both recorded revisions, and reproduce the
standalone-workspace release flags. Execute build/tests before sequential
measurement, with no competing workloads. The captured binaries use Rust/Cargo
1.98.1, debug level 1, frame pointers and unwind tables, pinned to CPU 2.

Seven attribution mutations additionally reject missing protocol, wrong capture
call graph, kernel events, zero periods, malformed headers, unparsed frames,
and fabricated source counters. Replay them with
`python3 docs/performance/results/change-0412/attribution/run-negative-probes.py --repo-root . --output /tmp/litchi-goal-0412-attribution-negative.json`.
