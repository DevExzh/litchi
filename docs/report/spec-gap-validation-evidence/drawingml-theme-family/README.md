# DrawingML `themeFamily` bounded profile

This directory contains the reproducible performance and allocation profile for
the shared `litchi_drawingml::theme::family` fragment API. It measures the
typed fragment owner in isolation with deterministic synthetic inputs:

* `small` is the native empty-element `thm15:themeFamily` fragment;
* `opaque` keeps the same typed attributes and adds bounded vendor attributes,
  comments, and an extension-list subtree that the typed projection does not
  interpret.

The matrix has four operation lanes for each fragment: `read`, a cheap
source-sharing `clone`, an exact semantic no-op commit, and `change`, which
edits the typed name. Each lane runs in three fresh processes with three warm-up
iterations and 30 measured samples per process. The process-local allocator
records allocation calls, direct allocation bytes, successful realloc old/new
sizes, requested bytes, peak-live deltas, and the source-sharing result.
Requested bytes charge direct allocation sizes plus each successful realloc’s
new size. Every sample requires the live-byte accounting identity, and startup
runs a raw alloc/realloc/dealloc counter self-test. No byte hashing or pointer
validation is performed inside a timed or allocation-counted closure.

Elapsed samples include the installed counting allocator observer and are
reported as allocator-instrumented operation time. Peak-live bytes are the
incremental peak above each timed closure’s live-before baseline. `/usr/bin/time`
RSS is whole-process startup/warm-up/sample RSS and is not attributed to an
individual operation.

This is absolute evidence for two synthetic fragments. It does not compare an
older implementation, establish a package-level theme edit speedup, or claim
whole-library throughput, host placement support, native Office acceptance,
or a general opaque-XML preservation guarantee beyond the exercised fixture.

The harness is deliberately outside the production workspace dependency graph.
Build and run it with `run_profile.sh`; the script records the exact command,
toolchain, host, binary SHA-256, pre/post Cargo package source manifests, dirty
theme hashes, and compact per-process JSON results under `results/`. The
pre-build and post-run manifests must compare byte-for-byte or the run fails.
Set `PERF=1` to request one `perf stat` capture for the opaque changed lane;
when the kernel denies counters, the script records the failure and retains
the allocator/timing evidence.
