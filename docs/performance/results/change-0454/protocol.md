# Change 0454 measurement protocol

This bundle measures the unnamed-slide PPTX cross-copy path after the focused
lossless relationship append work. The preserved lifecycle binary is bound by
its source manifest and hash in `baseline-lifecycle-binary.json`; the current
candidate binary and external probe are supplied by `build.py` in
`candidate-build.json`. A revision label alone is not used as build identity.

`protocol.json` is frozen before any lane is captured. The provider control is
the existing `pptx_provider_lifecycle_v1` report over the unchanged plain and
media-rich synthetic corpora. It retains 16 fresh child processes, each with
three warmups and 30 samples, pinned to CPU 2 with one worker. The lane order
is the change-0453 forward matrix followed by its exact reverse matrix:
bytes/range × plain/media-rich, with baseline and candidate in each matrix.
This gives two directional observations for every provider/corpus pair while
keeping each individual p50, p95 and p99 visible.

The pinned LibreOffice QA `smoketest.pptx` fixture has one unnamed slide and a
noncanonical relationship part. The candidate external probe runs two
descriptive lanes, one direct-bytes and one 256-byte/100-microsecond logical
range adapter, with the same three warmups and 30 retained samples. The
preserved baseline has a typed unnamed-slide refusal in the historical
name-only outcome record, so this fixture has no baseline timing population.
Its timings, output bytes, lexical prefix/suffix checks, exact copied payloads,
semantic reopen and eager reopen are evidence of the accepted candidate path;
they carry no speedup claim against a refusal.

`capture.py` wraps each workload in a fresh `taskset` child and GNU
`/usr/bin/time -v`. The Rust report remains the authority for API phase clocks,
logical read counters, budget work and preservation/output identities. The
resource log contributes process-level maximum RSS, page faults, filesystem
counter observations, context switches and user/system time. It includes
fixture setup and the retained sample loop, so it is descriptive process RSS,
not an operation-local peak or a physical I/O measurement.

`derive.py` computes every individual phase and API p50, p95 and p99 using
linear interpolation and computes a deterministic 10,000-resample bootstrap
95% interval for each lane median. It retains all paired timing and process
resource changes above 5% as review flags, and reports provider logical
read/work mismatches explicitly. The provider comparison is descriptive
control evidence; the external comparison is explicitly withheld.

The ordinary binaries do not expose operation-scoped allocator counters. This
bundle therefore reports allocation calls, allocated/deallocated bytes, live
bytes and allocator-region peaks as unavailable. RSS is retained as a separate
process observation and is never used as an allocation substitute. No cold
cache, native Office application, physical network/storage, scaling or
allocator claim is authorized.

Reproduction after the coordinator has released the serialized CPU gate:

```sh
python3 -B docs/performance/results/change-0454/machine.py
python3 -B docs/performance/results/change-0454/capture.py --suite provider --lane 0
# Capture provider lanes 1 through 15 in protocol order.
python3 -B docs/performance/results/change-0454/capture.py --suite external --lane 0
python3 -B docs/performance/results/change-0454/capture.py --suite external --lane 1
python3 -B docs/performance/results/change-0454/derive.py
```

The coordinator's bundle verifier must validate every receipt, binary/source
binding, report oracle, resource artifact and derived hash before cleanup.
