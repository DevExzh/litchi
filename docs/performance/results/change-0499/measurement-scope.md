# 0499 measurement scope

The 0499 comparison measures one selective OPC workload: ordered reads of
stored `Part` payloads through the source backed package API. The unchanged
harness constructs two synthetic archives: four 1 MiB stored Parts in the
`few-large` corpus and sixty-four 16 KiB stored Parts selected from a 512-Part
`many-small` corpus. It exercises owned bytes, a warm file source, and a fixed
short-delay instrumented source. Each case runs serial reads and batch widths
1, 2, 4, and 8, with three warmups, two repeats, and thirty timed samples per
repeat. The six warmup rows remain in every raw CSV; comparison statistics use
the sixty measured rows after the warmup boundary.

Elapsed values are the harness's operation-local read timings. Child-level
maximum RSS and the `perf stat` counters include package setup, fixture
construction, warmups, timed reads, output, and verification. The clone traces
are whole-child diagnostic counts; their traced elapsed times are excluded
from latency evidence. CPU affinity is 0–7 for measurements and 16–31 for
gates on a shared host. This bounds CPU placement but does not isolate memory,
filesystem, or other host activity.

The sources are warm local inputs or a deterministic synthetic short-delay
provider. They do not establish cold-cache behavior, remote-service behavior,
or production network performance. The corpus is a focused Part-read probe,
not an OOXML CRUD or end-to-end document workload, so these measurements do
not support broad CRUD throughput claims.

The loader cleanup assertion and panic recovery test describe unwind builds.
The workspace's ordinary release profile uses `panic = "abort"`; the evidence
makes no recovery claim for an aborting process.
