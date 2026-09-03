# Change 0393 PPTX selected-image evidence

This directory retains the matched release `A1/B1/B2/A2` evidence for the
PPTX selected-image query optimization. Each leg used one worker, five
warmups, 30 retained samples, release mode, and the allocator-enabled harness
under explicitly selected Rust/Cargo 1.98.1.

The four raw reports are zstd-compressed with their JSON bytes unchanged.
`abba-summary.json` is the profiler's immutable pre-adjudication capture
summary. `allocation-metrics.json` retains all operation allocator samples and
`abba-metrics.tsv` retains the percentile/delta projection. The parent
decision is recorded separately in `adjudication.json`; it accepts only the
selected `image` p50 result and exact selected-path allocator reductions.

`performance_claim: scoped`; `claim_authorized: true`.

The raw and compressed report hashes are bound in `evidence-manifest.json`.
To check integrity and inspect a report without retaining a decompressed file:

```sh
zstd -t docs/performance/results/change-0393/a1.json.zst
zstd -q -d -c docs/performance/results/change-0393/a1.json.zst \
  | sha256sum
zstd -q -d -c docs/performance/results/change-0393/a1.json.zst \
  | jq -e . >/dev/null
```

The expected raw A1 SHA-256 is
`cd8c7dafe0c8848abd778a2cb76b6259544c10e11622ce4f754dcbe0405c5920`.
