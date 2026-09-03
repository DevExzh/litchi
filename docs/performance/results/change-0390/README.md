# Change 0390 allocator reports

This directory retains the six release operation-allocator reports and their
six schema-v2 corpus catalogs for the one-decoder-session OPC materialization
change. The `control-*` and `candidate-*` report pairs cover the deterministic
`tiny-compressible`, `many-small-incompressible`, and
`few-large-incompressible` corpora. Each report has three warmups and 15
retained samples.

The reports and catalogs are zstd-compressed at level 3 with the underlying
JSON unchanged. The [compact summary](../opc-source-materialization-decoder-session-0390-summary.json)
binds the baseline source revision, candidate patch, source-file blobs,
compiler, both allocator binary hashes, corpus hashes, vector bounds, and
compressed/decompressed artifact hashes. The [checked comparison](../opc-source-materialization-decoder-session-0390-comparison.json)
is rederived from the six report members in memory with the existing
operation-scoped allocator policy.

The mechanism is one reusable Deflate decoder session per unmanaged full
materialization; stored entries bypass it. Relative to a fresh decoder for
each entry, the measured candidate delta is -2 calls/-80,320 bytes per avoided
decoder. The three corpus deltas are 4 calls/160,640 bytes,
510 calls/20,481,600 bytes, and 6 calls/240,960 bytes. Logical source reads,
returned bytes, and materialized Part counts are invariant between control and
candidate.

`performance_claim: none`; `claim_authorized: false`. Latency, operation-local
peak/RSS, copied/decompressed/physical-I/O counters, and broad or default OPC
behavior are withheld.

To inspect and validate a report/catalog pair without checking in temporary
files:

```sh
zstd -q -d -c \
  docs/performance/results/change-0390/control-tiny-compressible.json.zst \
  > /tmp/control-0390.json
zstd -q -d -c \
  docs/performance/results/change-0390/control-tiny-compressible.corpus.json.zst \
  > /tmp/control-0390.corpus.json
python3 tools/validate_perf_corpus_binding.py \
  --report /tmp/control-0390.json \
  --catalog /tmp/control-0390.corpus.json
```

The temporary paths are examples and are not part of the retained evidence.
