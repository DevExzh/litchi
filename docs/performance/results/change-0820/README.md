# 0820 — allocator test quality repair

The planned real-file durability experiment stopped at its original test gate.
The process-wide allocator test incorrectly required one allocation to increase
net live bytes despite unrelated allocator callbacks. The bounded repair changes
only `#[cfg(test)]` assertions to check signed byte conservation and monotonic
high-water accounting. Runtime allocator and production code are unchanged.

The original frozen plan, source witnesses, and failed quality log remain intact.
`repair/` records a separate source origin, focused regression, and fresh six-gate
quality run. See [the report](../../0820-allocator-test-quality-repair.md) and
[repair review](repair-review.md). Completion is recorded in `repair/quality.json`.

There are zero release builds, exports, qualification reports, and timing samples.
The frozen plan's 216 reports / 4,488 samples are planned counts only. Its 24
format/phase/policy cases remain queued for a fresh packet. The unexecuted
success-reader drafts are archived in `planned-readers/`; the current validator
checks the failure and repair evidence. The admitted 0819 baseline is unchanged.

After review and cleanup, replay with:

```sh
python3 -B docs/performance/results/change-0820/validate.py --final
python3 -B docs/performance/results/change-0820/seal.py --check-head
```

No iWork work is included. Unrelated workspace changes remain untouched.
