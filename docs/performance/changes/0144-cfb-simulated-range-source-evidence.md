# Change 0144: simulated CFB range-source selective-read evidence

Date: 2026-08-16

## Scope

The standalone performance harness adds six opt-in controls over the existing
deterministic final-position CFB selective-read corpora:

- `cfb_selective_simulated_mini_legacy_read`
- `cfb_selective_simulated_mini_shared_read`
- `cfb_selective_simulated_mini_4095_legacy_read`
- `cfb_selective_simulated_mini_4095_shared_read`
- `cfb_selective_simulated_fat_legacy_read`
- `cfb_selective_simulated_fat_shared_read`

The controls emit only the existing `many-small` and `wide-root` shapes. They
reuse the 36-byte MiniFAT, 4095-byte MiniFAT-boundary, and 4 MiB FAT targets;
corpus generation and deterministic validation remain outside the timed
interval. The legacy control uses a harness-only bounded delayed `Read + Seek`
adapter. The shared control uses the existing positional `ReadAt` simulator.
Neither changes a production crate or public reader API.

## Recorded evidence

Each measured sample retains stage-local open/read timings, their checked sum,
source counters, returned target length, and target SHA-256. Adapter creation
and evidence snapshots remain outside those timings. An additional
simulation object records, separately for open, selected read, and their
combined total:

- logical request count and returned bytes;
- physical request count, returned bytes, sorted requested sizes, and fixed size
  buckets;
- the configured simulation parameters; and
- a deterministic simulated-service floor obtained by summing the configured
  fixed latency, request overhead, and transfer time for every physical
  request.

The focused harness tests require every physical request to be non-zero and at
most the configured maximum, require count/size/bucket arithmetic to agree,
require the shared read stage to perform exactly the selected target work,
require MiniFAT legacy work to exceed the shared target work, and retain the
matched FAT read-work control. Target hashes and phase counters must remain
identical to the deterministic corpus. Separate short-read, EOF, and empty-read
tests keep requested-size and returned-byte accounting honest.

## Claim boundary

This change supplies reproducible simulated-source evidence only. The release
record accepts p50/p95 only for the named configured simulator. It is not a
cold-filesystem, physical-device, or ambient-network measurement and does not
accept production-source scheduling, p99, allocation, RSS, or native
DOC/XLS/PPT performance. A configured service floor is a model accounting
quantity, not observed wall-clock service.

The default matrix remains 36 cases / 198 result records. The six names raise
the selectable case-name count from 265 to 271. iWork remains deferred.

Focused smoke:

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  simulated_selective_cfb -- --nocapture
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 --shape many-small,wide-root \
  --case cfb_selective_simulated_mini_legacy_read,cfb_selective_simulated_mini_shared_read,\
cfb_selective_simulated_mini_4095_legacy_read,cfb_selective_simulated_mini_4095_shared_read,\
cfb_selective_simulated_fat_legacy_read,cfb_selective_simulated_fat_shared_read \
  --json target/perf/cfb-selective-simulated-read.json
```

Verification on the implementation revision:

- full standalone harness unit suite: 113/113, including the short-read, EOF,
  empty-read, many-small, and wide-root regressions;
- warning- and deprecation-denied all-target Clippy;
- rustfmt and diff checks; and
- two independent read-only reviews, including correction of asymmetric timed
  source construction and short/EOF request accounting.

## Clean release evidence

The implementation was committed as `c9f755026423d8f1a4771413be8461a6d8f40b49`.
One exact release binary (SHA-256
`60658045f3278a735928a4818289f55362caed22b2435d7e34600c0eeb4f3f51`)
ran from a clean detached worktree on CPU 2 in `A1 legacy, B1 shared,
B2 shared, A2 legacy` order. Every target/shape result has 20 warmups and 200
samples. The fixed model was 100 us latency plus 25 us overhead per request,
50 MiB/s bandwidth, and a 64 KiB physical-request ceiling.

| Target / shape | Selective read work, legacy -> shared | Total p50 reduction, pair 1 / pair 2 | Total p95 reduction, pair 1 / pair 2 |
|---|---:|---:|---:|
| 36-byte MiniFAT / many-small | 4 requests / 261,184 B -> 1 / 36 B | 40.12% / 39.99% | 40.64% / 39.08% |
| 4095-byte MiniFAT / many-small | 5 requests / 265,216 B -> 1 / 4,095 B | 40.09% / 39.82% | 40.26% / 39.75% |
| 36-byte MiniFAT / wide-root | 32 requests / 2,096,192 B -> 1 / 36 B | 41.96% / 41.83% | 42.23% / 41.58% |
| 4095-byte MiniFAT / wide-root | 33 requests / 2,100,224 B -> 1 / 4,095 B | 42.00% / 41.84% | 41.96% / 41.70% |

The 4 MiB FAT controls retain exactly 64 requests, 4,194,304 returned bytes,
and an 88,000,000 ns modeled read-service floor for both implementations.
Their paired p50 changes range from -0.09% to +0.08%, so they are classified
as matched-work near-neutral controls. Corpus and selected-target hashes,
returned lengths, configurations, revision, profile, affinity, sample counts,
and selector-paired identities match across all four legs. Every leg's request
buckets and service-floor arithmetic validate against its own observations.

The [compact summary](../results/cfb-simulated-range-0144-summary.json) retains
the exact comparison values and artifact hashes. The complete per-sample raw
records are committed as compressed
[`A1`](../results/cfb-simulated-range-0144-a1.json.zst),
[`B1`](../results/cfb-simulated-range-0144-b1.json.zst),
[`B2`](../results/cfb-simulated-range-0144-b2.json.zst), and
[`A2`](../results/cfb-simulated-range-0144-a2.json.zst) JSON.
