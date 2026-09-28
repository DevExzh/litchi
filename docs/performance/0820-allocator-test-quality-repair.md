# 0820 — allocator test respects process-wide live memory

The allocator wrapper test now checks process-wide byte conservation instead
of assuming that one allocation must increase the global live-byte total by
its size. A fresh quality run exposed the invalid assertion before the planned
durability measurements could begin. Only the `#[cfg(test)]` section of
`tools/perf-baseline/src/bin/support/counting_allocator.rs` changes; the
allocator implementation, counter implementation, production crates, and
benchmark runtime remain unchanged. This batch has **zero release builds,
exports, qualification reports, or timing samples**. The 0819 baseline remains
the current admitted real-file measurement.

[Original frozen plan](results/change-0820/plan.json),
[failed quality evidence](results/change-0820/quality-0/failure.json),
[repair origin](results/change-0820/repair/origin.json),
[independent repair review](results/change-0820/repair-review.md), and
[final results review](results/change-0820/results-review.md).

## Failure and repair

At base `096b810f23cc66fe01ec8c96c74f36888f954eff`, formatting and compilation
passed. The full all-feature test invocation, using two test threads, then
stopped in `docx_replayable_tail_append`. The first failing assertion was:

```text
during.live_bytes >= before.live_bytes + 256
```

The run had reached six test-result summaries: 568 passed, two failed, and one
ignored. The second failure was the next test receiving a poisoned test mutex
following the first panic. The complete failure log, original source hashes,
commands, environment, and frozen input witnesses remain unmodified.

`TEST_LOCK` serializes the tests' direct allocator operations, but the allocator
also observes callbacks from the test harness and other threads. An unrelated
deallocation between snapshots can outweigh the test's successful allocation.
The observed assertion does not prove that the allocator failed to record the
allocation; its required net increase is not a valid process-wide invariant.

The revised test retains its direct allocation/deallocation, successful pointer,
allocation-call, allocated-byte, deallocation-call, and deallocated-byte checks.
It additionally checks the following signed identity both after allocation and
after deallocation:

```text
live_after - live_before
    == (allocated_after - allocated_before)
       - (deallocated_after - deallocated_before)
```

The calculation uses `i128` so a legitimate negative net live-byte change is
representable. It also checks that the high-water counter is monotonic and
covers the current live-byte total. Counter callbacks and snapshots use the
same observer mutex, making each snapshot internally coherent. Existing
isolated `Counters` tests retain exact live/peak expectations for allocation,
deallocation, reallocation, failure, multi-step accounting, and overlapping
concurrent allocations. No runtime counter behavior or unsafe allocator code
changes, and the test-thread count remains two.

The retained failed execution is the before evidence. The focused wrapper
suite passes after the repair. It does not identify which unrelated callback
interleaved with the original assertion or claim a measured flake probability.

## Verification and scope

The focused allocator binary passes all five tests. The fresh full all-feature
run passes **641 tests, zero failures, and one ignored** across 28 result
summaries, including the doctest invocation. All six gates pass: formatting,
all-target compilation, full tests, Clippy with warnings denied, rustdoc with
warnings denied, and crate-boundary validation. The latter checks 65 packages
and 244 dependency edges while retaining 11 existing debt entries.
[Quality receipts](results/change-0820/repair/quality.json) and
[raw-log test summary](results/change-0820/repair/test-summary.json) bind these
results independently of the original failed invocation.

[Repair source](results/change-0820/repair/source.json) binds the current
9,197 production files and 87 harness files. Exactly one harness file differs;
its bytes before `#[cfg(test)]` equal the archived original. Both dependency
locks, all 35 normative input hashes, and the three unrelated workspace-file
hashes remain unchanged. The original experiment's frozen plan still records
its unchanged-source hypothesis; the separate repair origin explicitly records
the subsequent test-only change, rather than rewriting the failed attempt.

All execution is root-owned and serial. No release or workload capture was
started after quality failed. Prepared success readers are archived under
[planned-readers](results/change-0820/planned-readers/) as unexecuted drafts;
they are not result evidence. The current packet validator instead requires
the original failure, completed repair gates, exact source boundaries, and
absence of downstream measurement output.

The original plan crossed three real files, two path-save phases, and four
policies (default/full/file-only/no-sync). Those 24 cases remain queued for a
fresh packet at the repaired base. Explicit full/default is the API-route
control; weaker policies change crash guarantees and cannot justify weakening
the default. No policy latency, allocation comparison, throughput, or speedup
is inferred in this repair batch.

## Replay and cleanup

After final review, cleanup removed the owned test-build target: 5,204 files
and 3,685,472,529 logical bytes. Scratch was never created. [Cleanup](results/change-0820/cleanup.json) retains deletion accounting;
[seal](results/change-0820/seal.json) binds the exact committed change set,
including the test-only repair. Prior performance packets and unrelated work
remain untouched.

```sh
python3 -B docs/performance/results/change-0820/validate.py --final
python3 -B docs/performance/results/change-0820/seal.py --check-head
```

Receipt descriptors retain absolute workspace paths. A relocated checkout
requires explicit path rebasing; retained verification is not a claim that
another machine reproduced the failed interleaving or any timing result.
