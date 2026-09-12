# 0527: reject the XLSX row-primary-arena pilot

The row-owned primary-span candidate failed its frozen end-to-end admission
gate. Production was restored to the accepted 0525 implementation. Four
baseline-compatible regression tests remain; the two arena-only tests are
retained in the candidate evidence but removed from the working source.

The candidate moved the same primary XML spans from individual cell boxes to
row-owned arrays with a range per cell. Both ordinary and provenance writers
used checked slices. Full XML validation, independent readback, source and
execution fences, XML limits, unknown markup and formula handling remained
unchanged. This was a private allocation/layout change, not a public API change.

The fresh ABBA campaign measured 2,440 native durations and 40 allocation
samples, pinned to CPU 2 with Rust 1.95.0 release builds. Each primary
shape/repeat used 200 samples after 20 warmups. Positive values below are
candidate reductions from the matched baseline; they are rejected-candidate
diagnostics, not retained production speedups.

| Shape / repeat | Total p50 | Total mean | Commit p50 | Allocation calls |
| --- | ---: | ---: | ---: | ---: |
| medium / 1 | 3.5854% | 3.7546% | 6.9909% | 18.9078% |
| dense-sparse / 1 | 3.3680% | 3.3550% | 7.8422% | 19.8814% |
| medium / 2 | 2.3329% | 2.4847% | 8.0431% | 18.9078% |
| dense-sparse / 2 | **1.8214%** | **1.8105%** | 5.7713% | 19.8814% |

Every shape/repeat required at least 2% total p50 and mean reduction, 5%
commit p50 reduction, and 8% allocation-call reduction. Dense repeat 2 missed
both total gates, so no conditional profiling, hardware counters or eager
captures were admitted. Those lanes are unmeasured, not passing. The threshold
was not relaxed and the failed repeat was not replaced.

Allocated bytes fell 1.5912% for medium and 4.6535% for dense-sparse, while
incremental region peak live bytes rose 0.0328% and 0.0458%, respectively.
Reallocations fell 76.4869% and 86.3760%; they were reported separately from
allocation calls. All 47 matched adverse flags above 5% and 71 same-build drift
flags remain in the evidence. No blanket tail-latency, RSS or stability claim
is supported. Row ownership bounds a growth window, not the lifetime of all
retained span data.

See the [evidence bundle](../results/change-0527/README.md) for frozen source,
commands, raw reports, receipts, exact replay, individual flags, final quality
results and cleanup verification. The candidate passed 1,294 XLSX tests before
measurement; the restored final source independently passed all 12 quality
gates with 1,297 successful test executions.

OLE2/OOXML performance remains the priority. ODF stays deferred and iWork is
excluded. Next work should quantify another avoidable scanner cost, such as
unused end-event namespace resolution, before proposing a fresh independently
measured change. This rejected layout change must not be silently reintroduced
or treated as an accepted enabler. The overall optimization goal remains open.
