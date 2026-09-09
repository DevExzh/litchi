# 0487 OPC replay buffering evidence

This bundle measures retaining parser-consumed replay bytes across short reads
inside the existing OPC adapter allocation. The [hypothesis](hypothesis.md)
records the pre-change mechanism and ADR obligations; the
[implementation review](implementation-review.md) records the contract decision
and failure-boundary coverage. No iWork production files belong to this batch.

The [results review](results-review.md) records measured improvements, every
adverse row above 5%, validation, and retained failures.

The [methods](methods.md) retain the same 18 arms as 0485: three workloads,
four deterministic input modes, two explicit replay-store routes, normal and
allocator builds, two process repeats, and 30 measured samples per child.
The accepted before attempt is `formal2`, using the retained 0485 executables.
The after attempt is `formal1`, built from this batch. All 144 formal children
remain separate observations; pilots are omitted and no control arm is removed.

The frozen comparison imports canonical 0484 report/fixture validators. Capture
receipts bind binaries, build records, source manifests, exact arguments,
environment, and output oracles. The new protocol-hash field identifies the
actual protocol file; the first failed capture attempt and its historical
helpers remain in `development/protocol1` and `validation`.

Use the recorded gate commands to reproduce captures in a fresh evidence
directory. Existing attempts and summaries are never overwritten:

```sh
python3 -B docs/performance/results/change-0487/compare.py analyze \
  --before-attempt formal2 --after-attempt formal1
python3 -B docs/performance/results/change-0487/summarize_profiles.py profiles1
python3 -B docs/performance/results/change-0487/verify_seal.py verify
```

The first two commands create exclusive summary outputs; after sealing, use
the verifier to check the retained inventory. Separate perf-stat/strace children
provide diagnostic counters and compare syscall counts with the retained 0485
after profiles. They include setup and oracle work and are not operation-only
latency measurements.

The full non-iWork performance goal remains open. This route does not establish
cold-cache behavior, concurrent scaling, native Office certification, atomic
save, or the complete CRUD and provider-intersection matrix.
