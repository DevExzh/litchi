# Change 0435 evidence bundle

This bundle compares fresh ODT paragraph creation with the existing buffered
Builder and a new bounded sequential provider. Read the
[change record](../../changes/0435-odt-bounded-plain-paragraphs.md) for the full
tradeoff: at 32,768 paragraphs, operation allocator peak falls from 22.45 MB
to 0.42 MB, while normal p50 latency is 3.35–3.37 times the candidate Builder's.
Whole-process RSS is essentially unchanged. Every regression flag is retained.
The retention claim concerns provider-added storage with a lazy source and
discard sink; callers can independently retain input or output.

## Inputs and scope

`protocol.json` freezes 64/8,192/32,768 paragraphs, normal/allocator modes,
30 samples plus three warmups, two repeats, and roles before-buffered,
after-buffered, after-streaming. Phases A1/B1/C1/C2/B2/A2 retain 36 reports and
1,080 samples. Six large normal perf stat/record captures are under `profiles/`.
All formal captures use the candidate checkout as ambient state; `before/` and
`after/` separately bind each executable to its build sources and hashes.

`summary.json` rederives per-report intervals, allocator vectors/live deltas,
paragraph throughput, RSS scopes, 38 matched review flags, and 10 repeat flags.
`formal-profile-summary.json` retains counters and stacks with whole-process
scope. These scopes include setup/hashing/oracle work for profiles and process
RSS, while the operation timer excludes them. `claims[]` stays empty.

`pilots/`, `buffered-hypothesis.json`, and `profiles/preparatory-before-buffered/`
retain the preparation separately. `checks/` retains terminal passing and failed
attempts. The copied strict-debt baseline and both comparison scripts preserve
29 pre-existing harness diagnostics. `sources/` deduplicates source manifests;
`versions/` retains earlier oracle/protocol/profile drivers used by preparation.
No fresh native runtime was available. Functional fixture coverage is not native
application interoperability proof.

## Portable verification

Copy this entire directory, including its inventory, then run from any directory:

```sh
python3 -B /path/to/copy/verify.py --portable-check --require-inventory --stage final
python3 -B /path/to/copy/lifecycle.py --stage final
python3 -B /path/to/copy/summary.py --check
python3 -B /path/to/copy/profile-summary.py --check
```

These checks need no Git checkout or retained binaries. `seal.py` deterministically
compresses logs/profiles and inventories every retained file. `replay.py` runs
copied verification plus eight independent mutations, writes its terminal receipt
only after the child exits, and requires another seal to include that receipt.
The semantic mutation calls the copied per-report oracle directly so artifact
hash rejection cannot mask semantic validation. Do not edit a sealed bundle
while its replay is running.

`cleanup.py` requires passing precleanup portable proof and removes only five
batch-owned `/tmp/litchi-goal-0435-*` directories. It preserves the workspace
and harness target directories and the user's pinned `docs/GOAL.md`. The final
cleanup and portable receipts record the actual outcome. Rebuilding workloads
requires the named Git revisions and the exact build/capture commands retained
in receipts; portable verification checks existing evidence without rerunning
benchmarks.

## Completed lifecycle

Precleanup and post-cleanup copied verification both passed, each with all eight
mutation probes rejected and restored successfully. Cleanup removed exactly five
scratch directories containing 1,823,520,117 regular-file bytes; both shared
target directories and the pinned user goal file were preserved. The recorded
post-cleanup proof requires no retained benchmark binaries.
