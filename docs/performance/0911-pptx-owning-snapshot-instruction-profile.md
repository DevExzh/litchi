# 0911 — PPTX owning cold snapshot instruction attribution

This batch profiles the initial opened-presentation transaction on the retained
0910 executable, with owning input and materialized ReadAt as controls. It makes
no production change. The profile is scoped to the generated 200-slide,
64-text-edit fixture; guest instructions are not native elapsed time.

## Scope and failed attempts

The selected owner is
`litchi_pptx::package::model::Package::opened_presentation_transaction`.
Its measured caller is
`litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`.
The owner includes catalog/slide/notes validation, MCE, package cloning,
complete-package fingerprinting, payload digests, and detached transaction
construction. Ingress, edits, publication, output digesting, semantic oracles,
and artifact writes lie outside the owner interval.

The first protocol incorrectly expected only one transaction call per process.
Fixture construction also calls the owner through
`build_pptx_source_edit_corpus`. V1 produced two positive dumps and was rejected.
V2 attempted to reset before and dump after the measured parent; its positive
dump still retained both owner contexts and was rejected. Both raw attempts,
commands, reviews, and failure receipts remain in the packet. No conclusion
about Callgrind's internal reset implementation is inferred from that failure.

The [v3 protocol](results/change-0911/protocol-review-v3.md) predeclares the two
contexts and collects a separate dump after each owner return. Every positive
dump must have exactly one owner call from an expected exact caller. Setup and
measured contexts are classified by caller, never dump order or cost. Both are
retained; only the measured context enters the route comparison. Each dump's
owner self plus immediate children and all-function self sum must independently
reconstruct its summary. The terminal dump must be zero.

## Admitted results

All four v3 children pass exact caller, raw-counter conservation, terminal,
source/output, semantic and sink gates. Each has one excluded setup owner
context and one measured owner context:

| route / repeat | measured snapshot Ir | excluded setup Ir |
| --- | ---: | ---: |
| materialized A | 1,000,666,812 | 126,538,787 |
| owned A | 1,099,868,747 | 126,681,314 |
| owned B | 1,099,912,679 | 126,577,157 |
| materialized B | 1,000,594,047 | 126,635,417 |

These are two descriptive instruction observations per route, not a native
speedup estimate or a statistical latency comparison. The comparable native
measurement remains [0910](0910-pptx-owning-input-native-control.md): owning
snapshot p50 15.865198 ms versus materialized 12.349500 ms, with the owning
route faster overall because publication is much cheaper.

## Attribution and next decision

The complete [raw reduction](results/change-0911/analysis.json) retains all
positive function self costs, direct owner children and nested target edges.
`opened_presentation_with_limits` accounts for 998.932–999.008 million Ir in
the materialized owner and 1,098.150–1,098.194 million in the owning owner.
Detached transaction construction is only 1.342–1.394 million Ir. These direct
children, the owner's 53 self instructions, snapshot clone/drop and its direct
167-instruction memcpy call reconstruct each measured owner exactly.

Nested complete-package fingerprinting costs 922.844–922.848 million Ir in
the materialized snapshots and 1,000.196–1,000.199 million in owning snapshots.
These inclusive nested values overlap the owner and are not additional costs.
Exclusive `sha2::sha256::soft::unroll::compress` costs exactly 922,100,520 Ir
in every measured dump: 92.149–92.155% of materialized guest instructions and
83.834–83.837% of owning guest instructions.

This software SHA result is a profiling limitation as well as an attribution.
The host exposes `sha_ni`, while the captured guest ran the software backend.
The raw profile does not prove the backend used during the 0910 native runs,
and software guest instruction share cannot estimate native SHA time or an
Amdahl speedup. Confirm native CPU dispatch and sample native snapshot costs
before selecting a SHA or fingerprint optimization.

Owning snapshots also show 50,868,707–50,874,914 exclusive memcpy Ir,
28,254,410–28,254,417 memset Ir, and 7,985,337 specialized CRC Ir. These are
instruction counts over the whole measured owner, not copied/decompressed
bytes or allocations. Source inspection links complete-package fingerprinting
to `get_part` for every Part; on a cold owning package this can force deferred
payloads. This is a bounded mechanism hypothesis, not causal proof of the
native snapshot difference or permission to weaken revision/conflict checks.
The [source assessment](results/change-0911/candidate-assessment.md) records
the exact deferred-payload edges, dispatch source and correctness boundaries.

## Reproducibility and limitations

The executable is the unchanged 0910 release build, 9,301,456 bytes, SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`.
A fresh offline replay reconstructs all 9,355 frozen source files and verifies
its build custody, native analysis, and raw ZIP censuses. No Rust build or
quality rerun is needed for this evidence-only change; prior formatting and
warning-denied Clippy results remain bound to the same source and binary.

Fresh profile children use Valgrind 3.26.0, CPU 12, zero warmups, one measured
operation, and the serial materialized/owned/owned/materialized schedule.
The binary has debug information disabled: exact function and caller names
support attribution, but source-line profiling is unavailable. Raw Valgrind
`brk segment overflow` warnings are retained; workload and raw scope gates
must pass independently. Instrumented durations are excluded from native claims.

The source is 17,017,139 bytes, SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`;
the output is 17,017,145 bytes, SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.
The fixture has 229 OPC Parts and 445 ZIP members, including eight deterministic
2 MiB media parts. The generated-fixture qualification does not establish
producer-wide support, cold-cache behavior, allocation/RSS, hardware counters,
concurrency/scaling, or physical codec/copy byte counts.

The complete non-iWork performance goal remains open. ODF is deferred; iWork
and its shared cherry-pick remain outside this batch. Target-branch integration
is pending that unrelated conflict resolution. Accepted ADRs are unchanged.

## Replay and cleanup

The [independent review](results/change-0911/result-review.md) records the
raw-scope checks and fail-closed analyzer verification.

[Offline replay](results/change-0911/replay.py) passes before and after cleanup.
It validates failed-attempt and admitted-command custody, reconstructs the
inherited native source, repeats its ZIP censuses/statistics, and reproduces
the final raw reduction byte-for-byte. The three marked capture working
directories were removed (195 logical bytes); no build tree or worktree was
created. All raw profiles and twelve generated source/output archives remain
locally available. Archives are explicitly Gitignored with tracked hashes;
the executable remains in the retained 0910 packet. A fresh clone must rebuild
the executable and regenerate archives before replaying their identities.

```sh
python3 -B docs/performance/results/change-0911/replay.py \
  --output /tmp/new-0911-replay.json
```

The [plan](results/change-0911/plan-v3.json) and
[capture helper](results/change-0911/capture_v3.py) retain exact commands,
environment, tool/binary identities and the declared two-context schedule.
Instrumented elapsed times remain diagnostics and are not a native benchmark.
