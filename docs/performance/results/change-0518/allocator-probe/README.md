# DOCX publication allocation probe

This directory contains an isolated allocator-instrumented copy of the
current `managed_paragraph_batch_perf` example.  It is evidence support for
the 0518 publication attribution work; it does not modify a production crate
or the normal benchmark example.

`generate.py` reads
`crates/litchi-docx/examples/managed_paragraph_batch_perf.rs`, requires its
exact source SHA-256 (`89ffdd564f6a76814fcc8b769f324557bd15a5201cb366ce9e75ade0a1c61181`),
and checks every transformation anchor occurs exactly once.  It writes the
auditable copy to `src/main.rs`, a unified `source-diff.patch`, and
`source-binding.json`.  The copied benchmark keeps the original fixture,
CLI, CSV schema, publication output, patch forward/inverse checks, managed
budget checks, source counters, cache gauges, output digest, untouched-member
payload checks, and all semantic output oracles.

The generated binary includes the canonical
`tools/perf-baseline/src/allocation_metrics.rs` and
`tools/perf-baseline/src/bin/support/counting_allocator.rs` through `path`
modules.  `extern crate self as litchi_perf_baseline` supplies the canonical
allocator wrapper's crate path.  The canonical wrapper is the only unsafe
code in this probe.  `allocation_metrics::enable()` is called at the start
of `main` when the `allocator-metrics` feature is enabled.

Each `run_sample` starts an allocation region immediately before
`Package::publish_document_commit_to_stream` and finishes it immediately
after that method returns, before the existing `drop(published)` call.  This
matches the publication method profile and excludes the caller's returned
`Snapshot` drop.  The existing lifecycle timer still includes the returned
snapshot drop, as in the normal CSV benchmark.  The measured allocation
sample is emitted as one JSON line with `"tag":"allocationSample"`, case,
repeat, ordinal, warmup, and an `allocationSample` object only after the
lifecycle clock and all output oracles have completed.  JSON printing is
therefore outside CSV timing and these are instrumented allocation timings,
not native timing rows.

The parent capture should use one child per case with
`--samples 1 --warmups 0 --repeats 1`; the JSON lines are in that child's
stdout and the unchanged CSV is written to the requested `--output` path.
The authoring step deliberately performs no build or capture.  The parent
driver owns the build receipt, executable custody, capture, and cleanup.
