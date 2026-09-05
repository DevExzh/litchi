# Change 0419 resource review

`review_scope: matched PPTX cross-copy media-rich and plain lifecycles`

`recommendation: retain the private writer change in its existing owned-source scope; make no general memory or latency claim`

## Result

All 16 matched ABBA journals are present in `summary.json`, with unchanged
corpus/output identities and same-revision drift below 5%. The allocator lane
uses 30 samples per leg and is the relevant resource comparison. In the
media-rich lifecycle, operation requested-allocation volume falls from
36,811,874,712.7 to 369,979,465.4 bytes (−98.995%), allocation calls fall
3.288%, and reallocation calls fall 21.763% in both pairings. The plain
lifecycle requested volume falls 36.217%; allocation calls fall 0.696% and
reallocation calls 5.410%.

Those requested-byte totals are allocator accounting: each successful
reallocation contributes its full requested new size. They do not measure
physical copying, simultaneously resident memory, or RSS. The large media
percentage therefore supports fewer or smaller allocator requests, not a
98.995% memory reduction.

The retained snapshots do not show a measured live-memory increase. Media
`live_bytes_after` is 287,885,879 for both revisions and both pairings; plain
is 884,426 for both. Media `peak_live_bytes_after` is 896,603,242 in A1/B1
and differs by only 1,568 bytes in A2/B2 (0.000175%); plain is 3,565,760 in
all legs. Allocator whole-process RSS is slightly lower for the candidate:
about −0.024%/−0.035% for media and −0.160%/−0.073% for plain. These small
RSS differences are descriptive whole-process observations, not evidence of
an operation-local memory saving.

The 100-sample normal lane is diagnostic and below the 500-sample release
latency threshold. Media p50 changes −1.595% and −1.675%, with p95/p99 also
lower. Plain p50 changes +2.071% and +0.380%; its p95/p99 directions are
mixed. Same-role drift remains below 5%, so this is not a release latency
claim. A final exact-size copy in `BoundedVecWriter::into_bytes` can coexist
with the grown buffer, and its copy/allocator cost is a plausible source of
a small plain-workload regression even though the measured snapshots are
unchanged.

## Evidence limits

The Heaptrack captures bind the control and candidate binaries to their
recorded revisions and the same protocol. They locate `BoundedVecWriter`
frames in the respective traces, but Heaptrack 1.5.0 fills its `-H`
allocation histogram before applying `--filter-bt-function`. Consequently
the retained filtered histogram is whole-command data in both analyses; its
requested-byte sum must not be presented as writer-only allocation volume or
as a control/candidate physical-copy comparison. The matched allocator
vectors in `summary.json` provide the scoped resource result.

The live and peak fields are snapshots of process allocator state. The peak
value is a process high-water observation and does not establish the peak
inside the timed operation. `/usr/bin/time` RSS likewise covers corpus
construction, correctness gates, setup and teardown. The experiment does not
exercise an archive near the default 128 MiB output limit, low-memory refusal,
or a workload where the final compaction buffer approaches that limit.

## Disposition

There is no resource-evidence blocker to retaining revision `0320a6a889` in
the already guarded owned-source path. The allocator result is strong and
the measured live/peak/RSS snapshots do not regress. The change should remain
scoped to its existing correctness, output-limit, typed-allocation-error and
exact-source checks. It should not be described as a memory improvement,
physical-copy reduction, or latency claim.

Before broadening the optimization, add an operation-local retained/peak
measurement and matched near-limit and low-memory cases. Verify that the
fallible final copy preserves the refusal/error contract and quantify its
transient coexistence with the grown buffer. Keep the existing fallback for
dirty/custom/non-default paths while that evidence is collected.

## Evidence

- [0419 summary](summary.json)
- [control Heaptrack analysis](heaptrack/control/analysis.json)
- [candidate Heaptrack analysis](heaptrack/candidate/analysis.json)
- [measurement protocol](measurement-protocol.json)
- [source review](source-review.md)
- [bundle README](README.md)
