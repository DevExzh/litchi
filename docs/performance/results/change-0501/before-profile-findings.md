# Before profile findings

The source-sensitive before capture is complete. The formal matrix has eight
serial provider-lifecycle children: plain and media-rich corpora, owned bytes
and recently written file-warm tmpfs lanes, three warmups, thirty measured
samples, and forward/reverse lane order. All eight reports passed the exact
payload/output oracles, semantic gates, source and cache counters, resource
budgets, whole-child resource capture, and cleanup checks. The retained sample
count is 240. The formal reports and receipts are under `before/`.

The supplementary profile uses media-rich owned and file-warm lanes with one
warmup and eight measured samples at 99 Hz with `dwarf,8192` callchains. The
exact invocation was:

```text
DEBUGINFOD_URLS='' PERF_BUILDID_DIR=/tmp/litchi-goal-0501/perf-buildid python3 -B docs/performance/results/change-0501/profile.py before
```

Both profile receipts passed, and each report's output, source archive, and
destination archive identities matched its paired `perf stat` run. Raw data
is retained in the owned tmpfs namespace:

* `/tmp/litchi-goal-0501/profiles/before/media-rich-owned/perf.data`
* `/tmp/litchi-goal-0501/profiles/before/media-rich-file-warm/perf.data`

The local `perf report` exports had 454 samples and resolved
`sha2::sha256::x86_sha::compress` at 31.01% for the owned lane and 30.08% for
the file-warm lane. These are whole-child samples including setup, corpus
gates, warmups, diagnostics, report construction, and serialization. The
callchains did not recover a reliable `digest_touched` or `digest_bytes`
caller, so the profile proves SHA-256 hotness in the child only; it does not
prove that the private touched digest is the sampled source of that hotness.
The raw perf data and compact exports remain available for a later candidate
comparison.

The first owned profile export was intentionally preserved after
`perf record` completed because its `perf report` process inherited
`DEBUGINFOD_URLS=https://debuginfod.ubuntu.com` and blocked on remote symbol
lookup. The owned report subprocess was terminated, its partial output was
retained, and the same raw perf data was exported again with
`DEBUGINFOD_URLS=''`. The recovery history, old receipts, and the pre-fix
verifier are in `interrupted/`; `driver-transition.json` records their hashes
and paths. A subsequent profile-driver retry reran the owned `perf stat` and
`perf record` workload, creating a distinct retained raw file, and then failed
only because the old verifier still required the formal 30/3 dimensions. The
final before invocation sampled both lanes again and passed. The profile
driver now clears `DEBUGINFOD_URLS` itself for the after phase.

The file lane is a warm recently written positional file in `/tmp` tmpfs. It
is a provider/control comparison and carries no cold-storage or physical-disk
claim. No optional range extension was run.
