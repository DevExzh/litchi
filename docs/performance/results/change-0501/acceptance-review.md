# 0501 acceptance review

The matched API measurements support retaining the private payload digest
reduction. Across two independently captured repeats, the media-rich owned
lane improves API p50 by 61.467–61.481% (2.595–2.596x), and the warm tmpfs file
lane by 51.452–51.870% (2.060–2.078x). Their p99 improvements are respectively
61.199–61.350% and 50.835–51.122%. The plain controls remain within 2% at API
p50. Plain warm-file R2 regresses 1.347% at p50 and 1.632% at p99; these raw
values are retained even though they do not cross the review threshold.

All 208 comparison rows are retained. The 52 absolute changes above 5% are
favorable media-rich timing/throughput results. No source archive, destination
archive, exact output, source-read histogram, cache, or resource-counter
fingerprint changes. Whole-child RSS changes stay within 0.042% for media-rich
lanes and range from -3.914% to +1.123% for plain lanes. This is no allocation,
physical-copy, decompression, or I/O reduction claim.

Each media-rich fixture contains eight 2 MiB image payloads. Code inspection
therefore identifies 32 MiB of redundant payload hash input eliminated across
the two preparation passes in the measured workflow. Exact payload comparison,
candidate reread, and compressed-entry authorization remain. The new helper
keeps the preceding hash traversal's 64 KiB cancellation checks.

The timing result is the sum of explicit API phase clocks, excluding fixture
construction, validation/oracles, diagnostics, serialization, and final drops.
Whole-child perf profiles contain those excluded activities and cannot supply
an Amdahl fraction for this API sum. Before SHA compression shares are 31.01%
owned and 30.08% warm-file; after shares are 27.95% and 26.93%. Incomplete caller
unwinding prevents precise fresh touched-digest attribution. The historical
0449 targeted shares are hypothesis evidence, not a causal prediction for
0501. Separate executable builds and non-interleaved phase ordering on a
shared host limit attribution of small differences; the reversed workload
orders are not ABBA interleaving.

The full default matrix is independently refreshed on the final candidate:
two 201-row repeats with 15 measured samples per row. Its validation binds the
generated catalog and report to the current CRUD index. It does not add native
notes/chart closure, promote correctness-only scenarios, establish cold
storage/concurrency coverage, or complete the overall non-iWork goal.

Final acceptance additionally requires the recorded crate gates and independent
verification of source, binary, helper, report, and profile custody. Failed
fixture/lint gates and supplementary profile attempts remain part of the
record; they are not counted as successful normal measurements.
