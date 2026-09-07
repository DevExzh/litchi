# 0455: larger bounded unchanged ZIP transfers

The ZIP preservation writer uses a 64 KiB stack buffer for unchanged local-member
spans, up from 32 KiB. This removes requests when the caller's range provider can
return 64 KiB. Preparation, layout validation, short-read handling, partial-write
accounting and owner-level source/resource/cancellation checks remain in place.
The larger buffer adds 32 KiB of stack per active publication; it adds no heap
allocation or worker. The new regression crosses both chunk boundaries with a
large stored member, short reads and a partial sink failure. The existing sparse
ZIP64 test now admits writes up to 64 KiB.

The PPTX lifecycle benchmark now retains its existing sink summary after timing
and allocation measurement finish. The matched source manifests differ only in
the production buffer constant and ZIP64 test bound. Both builds contain the
same regression test and sink instrumentation.

In the 720-sample formal matrix, media-rich publication reads fall from 833 to
577 and sink writes from 1,023 to 767. Returned publication bytes remain
16,830,603 and output remains 33,599,715 bytes with the same SHA-256. The full
32 KiB request count changes from 512 to 256 full 64 KiB requests. Plain workloads
have unchanged request and write counts.

For the simulated range provider (64 KiB cap, 200 us per request, 25 MiB/s,
separate sleeps), API medians fall 4.775% and 4.449% across two repeats;
publication medians fall 9.227% and 8.606%. The nominal fixed delay reduction is
51.2 ms. Observed publication savings are about 78.5–84.8 ms; this evidence does
not separately attribute the remainder to sleep overhead or CPU work.

All 19 absolute >5% timing flags remain visible. Two positive flags affect the
first range/media-rich open tail: source-open p99 +28.241% and combined-open p99
+12.929%. Neither repeats in the second comparison. The initial bytes/media-rich
R2 result also has an unexplained 17.817% API decrease with a 19.674% planning
decrease, while R1 shows only 1.318% lower API latency. The unchanged mechanism
counters do not explain that discrepancy. A separately declared 240-sample
bytes/media confirmation retains all eight processes: three paired API changes
are -0.246%, -0.742% and -0.637%; the fourth is -27.962% with a slow
baseline process. These results do not establish a repeatable large bytes gain.

Operation allocation counts and bytes are unchanged. Media-rich publication
uses 5,397 allocation calls and 21,057,867 allocated bytes, with regional heap
peak 17,406,388 bytes above operation entry. The candidate's lower absolute live
counter begins before the operation and is not an operation allocation saving.
Process RSS comparisons stay within 2.671%; stack growth is outside allocator
counters. Separate `perf stat` runs record whole-process cycles, instructions,
branches, misses and page faults, including corpus generation and output oracles;
they are not API-attributed hardware counters.

The pinned LibreOffice QA self-pair publishes the exact previously verified
42,948-byte output through bytes and range providers. ZIP CRCs, member order,
copied slide/image bytes and XML parsing pass. This is not an independent native
document pair or a native-application roundtrip.

The original whole-process counter experiment is adverse: candidate mean
instructions rise 9.25% and cycles 12.00%, with about 1.217 million page faults
versus baseline 381 thousand. The candidate API medians in those instrumented
processes are about 35.34 ms versus baseline 25.28 ms. This observation is retained,
including all raw reports; the external Python oracle runs after perf exits.

A separately declared instruction-sampling pair with local symbols produces
about 48.65/48.60 billion sampled instructions and API medians 25.360/25.232 ms.
Compression of the corpus and SHA-256 dominate both profiles. Restricted kernel
symbols and incomplete unwinding prevent precise kernel attribution. The first
report attempt stalled with remote debug-symbol lookup enabled and was stopped;
its failed receipt, script and partial artifacts remain in the bundle.

Slow processes also occur in the original baseline: its ordinary R2 has
942,424 minor faults and 30.510 ms API median; the final confirmation baseline
has 1,221,860 faults and 34.827 ms. Fast processes of both builds have roughly
381–385 thousand faults and 25.0–25.4 ms medians.

An eight-process ABBA diagnostic then fixes the glibc mmap threshold to 128 KiB
and 32 MiB in turn. Explicitly setting this threshold disables glibc's dynamic
adjustment, as documented in the [glibc manual](https://sourceware.org/glibc/manual/latest/html_node/Malloc-Tunable-Parameters.html).
At 128 KiB, both builds produce about 1.779 million faults and 38.5–38.8 ms API
medians; candidate instructions change -0.057%/-0.097% and cycles
-0.171%/+0.057%. At 32 MiB, both API medians are about 25.2–25.4 ms; one baseline
still has extra process faults. All four controlled candidate API changes are
between -0.281% and -0.714%. This supports allocator paging as a major source of
process variability, but does not reconstruct the historical mmap decisions or
prove the exact cause of every original outlier. These are diagnostic runs,
separate from the ordinary and allocator-instrumented populations.

Retain the 64 KiB buffer for the repeatable range-request reduction and range
latency result. Do not claim a CPU improvement or a large bytes-only speedup.
The original adverse counters and process variability remain limitations.
The fixed stack cost is explicit and independent of document size.

Release checks pass 456 ZIP, 497 OPC, 854 PPTX and 381 harness tests (2,188 total,
6 ignored), strict lint, rustdoc, formatting, minimal workspace and boundaries.
The unchanged OPC sanitizer fuzz harness passes 1,000 runs with the retained
corpus, including large stored/deflated custom parts. Precleanup and separate-copy portable
replay pass; owned temporary artifacts are inventoried and removed.
The full non-iWork goal remains open: broader CRUD coverage, native applications,
physical cold I/O, bounded existing-document append/repackaging and representative
bounded-worker scaling still need evidence.
