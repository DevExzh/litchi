# 0447: Explicit transfer-rate pacing for caller-owned range sources

Previous turn: progress, committed 0446. No owned CPU job is live. The accepted
ADR tree remains c950b6c8be822561b498d7bbe87c460873dcbf49, unchanged from its prior
complete read. Production packages and iWork remain outside this tooling batch.

GOAL.md requires range simulation with configurable latency, bandwidth, request
overhead and range size. The existing adapter has a fixed request delay and
short-read cap but no transfer-rate control. Add an optional nonzero bytes/second
setting, preserving the existing configuration when absent. Each successfully
returned nonempty chunk requests ceil(bytes * 1e9 / rate) ns of transfer pacing,
after the existing pre-delegation fixed request delay. Track the exact requested
transfer-delay sum and count separately from actual elapsed time. OS sleep
rounding/oversleep is not physical network I/O or proof of achieved bandwidth.
Concurrent reads have per-call pacing; this is not a shared-link bandwidth cap.

The CLI exposes this only for the explicit range provider and bounds the rate
so accidental multi-hour sleeps are rejected before corpus construction. Empty,
zero-cap, EOF and failed reads must not fabricate transferred bytes. Checked
integer arithmetic and fail-closed counters remain. Sources and versions stay
caller owned. No ambient networking or production runtime behavior is added.

Run adapter boundary/short-read/failure/overflow tests, CLI misuse tests and exact
PPTX lifecycle equivalence. Then freeze a matched single-build plain/media-rich
PPTX matrix with 64 KiB range caps and 200 us fixed request overhead, unpaced and
25 MiB/s paced transfers, 30 samples/3 warmups, CPU 2, one worker, two reversed
repeats. Preserve every raw phase/counter/budget/output observation and review
absolute 5% repeat drift. This is a simulation baseline and measured enabler,
not a production improvement or real remote/cold/scaling result. Normal and
allocator metrics must not be conflated; this custom lifecycle lacks operation
allocator attribution and reports that limitation explicitly.
