# Change 0436 evidence bundle

This bundle measures private ordinary-text Work batching in fresh ODT
sequential publication. Read the [design](design.md) for the pre-implementation
hypothesis and the [change record](../../changes/0436-odt-bounded-text-spans.md)
for the measured outcome and limitations. Normal p50 is 19.382–30.156% lower
across the three sizes and two repeats. Exact bytes and aligned allocator
vectors match; operation peak remains 420,091 bytes, with no RSS improvement
claim. None of 78 matched comparisons crosses 5%; one tiny p99 repeat flag
(−5.489%) remains visible.

## Inputs and scopes

`protocol.json` pins the same `odt_streaming_create` API before and after,
64/8,192/32,768 paragraphs, normal and allocator binaries, 30 samples and three
warmups, two repeats in A1/B1/B2/A2 order, CPU 2 and one worker. The formal
matrix contains 24 reports and 720 samples. Four large normal perf stat/record
captures live under `profiles/{before,after}/{stat,record}`. Formal captures
share candidate ambient checkout state; build descriptors bind each retained
executable to its own build revision and source manifest.

Exact archive, content XML, styles, metadata, semantic projection, topology,
output hash, accepted sink bytes and sink calls must match. The copied 0435
per-report oracle validates both roles as `after-streaming`; the 0436 outer
verifier owns cross-revision identity and source custody. The inner oracle's
historical change number is deliberately unchanged.

The operation timer includes fresh paragraph construction, publication and
HashingDiscardSink writes. Corpus setup, package reopen, semantic gates and
sink digest extraction are outside that timer. Perf profiles and GNU time RSS
cover the whole process. Operation allocator peak above entry, allocation
traffic, live balance, process high-water vectors and GNU time RSS remain
separate metrics. `summary.json` rederives individual results, uncertainty,
regression/repeat flags and raw profile counters/self rows.

Six fresh before pilots and six candidate pilots remain separate from formal
measurements. `batching-hypothesis.json` binds the baseline pilots and retained
0435 raw profile: both baseline executable hashes exactly match that prior
candidate, and Cargo reused the cached build. `preparatory-summary.py --verify`
checks these retained inputs without requiring the prior bundle. Failed
attempts remain terminal receipts under `checks/`; source manifests are
deduplicated under `sources/`.

## Portable verification

Copy this complete directory, including its inventory, and run:

```sh
python3 -B /path/to/copy/verify.py --portable-check --require-inventory --stage final
python3 -B /path/to/copy/lifecycle.py --stage final
python3 -B /path/to/copy/summary.py --check
python3 -B /path/to/copy/preparatory-summary.py --verify
python3 -B /path/to/copy/decision.py --verify
```

These checks need no Git checkout or retained binaries. `seal.py` compresses
logs and raw profiles deterministically and inventories every retained file.
`replay.py` runs copied verification and eight independent mutation probes;
it writes its terminal receipt only after its child exits, then requires a
new seal. The semantic mutation invokes the copied oracle directly so an
outer artifact hash failure cannot mask the semantic gate. Keep the source
bundle unchanged while replay is running.

`cleanup.py` requires portable precleanup proof, removes only the three
`/tmp/litchi-goal-0436-{binaries,odt,odt-tests}` directories, and preserves both
shared Cargo target directories and the pinned user goal document. Rebuilding
workloads requires the named revisions and receipt commands; portable replay
verifies the retained measurements without rerunning benchmarks.

The [capture-time verifier](versions/README.md) is retained exactly. Current
replay corrects an inherited five-directory cleanup expectation to the actual
three paths and validates original capture/build hashes against that retained
driver. The full matrix passed both original and corrected verifiers before
sealing; measurements and receipts were not rewritten.

## Completed lifecycle

Precleanup and post-cleanup copied verification both passed, each rejecting
and restoring all eight independent mutations. Cleanup removed exactly three
scratch directories containing 1,822,597,957 regular-file bytes. Both shared
Cargo target directories and the pinned user goal were preserved. Portable
proof succeeds without the retained benchmark binaries or external drafts.
