# Change 0495 profile review

This is a descriptive review of the retained [profiles1 summary](profiles/profiles1/profile-summary.json), not an operation-level CPU attribution. Each observer child includes process startup, corpus and provider setup, warmups, the edit/commit/publication path, output hashing and semantic checks, and report serialization. The `perf stat` and `strace` values below therefore describe the complete child. Operation-local timing, allocation, budget, source, sink, and output claims remain the validated Rust report's responsibility.

The summary is schema `docx-edit-provider-managed-profile-v1`, phase `after`, normal role, source revision `de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8`, and CPU 2. It covers both APIs and the owned, file-warm, and short providers with three measured rows and one warmup per child. The retained binary is 478,225,064 bytes with SHA-256 `d94d8cac8bc1628148f54de67531f003ec2d1180b5a81bbc1b7c0c47dc13a823`; the profile summary SHA-256 is `dedfa9235a5fe110543e10888c1bc1bddb824ae6e3c4f0d6869b3a17a116eb56`.

All six cases have unchanged source manifests, successful private-TMP cleanup, and canonical report identity validation for the selected perf and strace runs. Perf's full event set was unavailable in every case because `LLC-loads` and `LLC-load-misses` were unsupported; the helper selected the core event set and marked each result `degraded`. Strace was available in all six cases. Its reported totals matched the parsed rows, with no unexplained rows. The two owned `perf record` runs produced valid 100-row reports and successful script exports: 1,238 samples for unmanaged and 1,262 for managed.

## Whole-child perf counters

The selected core counters were task-clock, cycles, instructions, branches, branch-misses, minor faults, and major faults. These are one observer run per case, so the table is a directional diagnostic rather than a statistical comparison.

| API | provider | task-clock (ms) | cycles (B) | instructions (B) | minor faults |
| --- | --- | ---: | ---: | ---: | ---: |
| unmanaged | owned | 1,107.29 | 4.969 | 10.782 | 108,837 |
| managed | owned | 1,094.65 | 4.917 | 10.746 | 88,269 |
| unmanaged | file-warm | 1,134.82 | 5.072 | 10.952 | 107,297 |
| managed | file-warm | 1,120.57 | 5.033 | 10.990 | 95,601 |
| unmanaged | short | 1,346.66 | 6.047 | 14.242 | 99,144 |
| managed | short | 1,351.96 | 6.074 | 14.366 | 105,379 |

The managed/unmanaged differences are small and change direction across providers. They do not establish a managed edit CPU improvement or regression because the counters include setup and validation, and there is no before-build profile in this receipt. The unsupported LLC events and zero-valued L1 fields in the full attempt are unavailable hardware observations; they are not evidence of zero cache misses.

## Syscall observations

Strace totals were 13,753 calls for both owned cases, 82,600 unmanaged versus 86,836 managed calls for file-warm, and 1,406,721 unmanaged versus 1,409,341 managed calls for short. Every case reported one syscall error; the parser retained it and the total-call identity check passed.

Owned unmanaged time was led by `brk` (44.05%), `read` (23.78%), and `write` (20.15%). Owned managed time was led by `brk` (39.15%), `munmap` (24.52%), and `read` (22.78%). File-warm time was led by `write` (72.72% unmanaged, 73.39% managed), with `statx` and `pread64` present in both. Short time was dominated by `write` (98.60% unmanaged, 99.04% managed). These are whole-child syscall shares: report and trace serialization, output writing, and provider I/O all remain in scope, so they cannot be assigned to the publication operation.

## Owned stack samples

The two call-graph exports are statistical samples of the whole owned child. No sample was classified as `run_sample_publish_ancestor` in either API (`run_sample_publish_period` was zero), so the capture provides no measured publication-stack share. That absence is a symbol/unwind classification result, not evidence that publication consumed no CPU.

| stack class | unmanaged: samples / all period / run-sample period | managed: samples / all period / run-sample period |
| --- | --- | --- |
| preflight | 17 / 1.243% / 1.979% | 15 / 1.110% / 1.737% |
| process setup or unclassified | 420 / 35.918% / 57.158% | 417 / 34.996% / 54.772% |
| run-sample output oracle | 756 / 58.839% / 93.634% | 757 / 57.603% / 90.154% |
| run-sample setup or teardown | 45 / 4.001% / 6.367% | 73 / 6.291% / 9.846% |

The largest classified leaf in the output-oracle class was `payload_bytes` in both captures. The process-setup class included `memmove`, deflate, and unknown or SIMD leaves. These are actionable places for a phase-separated follow-up, but they do not identify an operation duration or prove a cause for the managed/unmanaged differences. The stack summaries are [unmanaged](profiles/profiles1/owned-unmanaged-api/perf-stack-summary.json) and [managed](profiles/profiles1/owned-managed-api/perf-stack-summary.json).

## Follow-up interpretation

The verified [formal1 receipt](verification/formal1.json) passes custody checks, while its analysis explicitly leaves `claim_authorized` false and describes the result as provider evidence. Its allocator rows show unmanaged allocator calls changing from 22,859 before to 9,696 after, and allocated bytes from 5,721,334 to 1,623,696. The same analysis retains repeat or paired flags for several whole-child RSS changes around 7.5–12.5% and for some latency comparisons, including file-warm cases. Those observations are useful follow-up triggers; they do not establish that managed ownership caused either change.

A useful next measurement should separate the timed edit path from output hashing, trace/report serialization, and package setup, then repeat the affected provider arms under the same source and binary custody. The current profile can guide that separation through its output-oracle and setup classes, but it cannot supply the phase-local CPU number itself. Any future cache statement must use supported events or a separate validated observer rather than the unsupported LLC rows.

