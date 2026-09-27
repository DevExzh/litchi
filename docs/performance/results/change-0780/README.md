# 0780 static baseline MCE capabilities

Measurement complete: 120 native, 40 allocator, ten baseline qualification, twelve perf-stat and two Heaptrack processes. Six final quality gates pass (5,683 tests, 35 ignored). See the report and disposition for the adoption decision and the large-lifecycle regression; constructor-only measurements are diagnostic. All 20 native spread flags and four paired metric flags remain retained.

The standalone probe, source census, dependency locks, exact build/capture commands, executable identities, raw process reports and quality logs are retained here. Allocation requests are not physical copies or resident bytes. No cold-cache, concurrency, complete-CRUD or native Office interoperability claim follows. The owned build target was removed only after all four executable path/size/SHA identities were checked; cleanup.json supplies exact witnesses for replay.

Offline replay:

```sh
python3 -B docs/performance/results/change-0780/validate.py
python3 -B docs/performance/results/change-0780/tables.py --check
```

The validator requires the final seal. Its library API permits preseal audit; both verify a seal when present. The seal covers every retained packet file except itself. Observer analysis enforces immutable identities and remains stable after executable cleanup and worktree relocation.

Reproduction requires a fresh packet/output directory and owned target, never overwriting this capture. Start from the base in origin.json; restore the retained workspace and probe lockfiles and identical probe source. Adjust the owned target path in custody.py for the fresh run. Build `before`, run `capture.py qualification` and `observers.py heaptrack-before`, and decode the trace with the exact recorded command. Apply `candidate/applied-model.patch` (not the historical draft), run `quality.py`, and build `after`. Then run `capture.py native`, `capture.py allocation`, `observers.py heaptrack-after` and `observers.py perf` serially in the frozen order; decode the second trace with the recorded flags. No sample-dependent retry or resampling is part of the plan. Generate observer analysis with the pure observe.analyze API, write the disposition, and run analyze.py/tables.py before sealing. Paths in new receipts identify that new reproduction; retained receipts must never be rewritten to pretend it was this run.

The initial quality-0 failure was a test-only iterator pattern corrected before all final candidate gates/builds. Historical candidate/model.rs and model.patch retain that draft; applied-model.rs and applied-model.patch are authoritative. Failed history, all primary samples, profiler traces and replay scripts are retained.
