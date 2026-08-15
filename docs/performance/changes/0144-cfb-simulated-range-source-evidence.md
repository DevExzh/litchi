# Change 0144: simulated CFB range-source selective-read evidence

Date: 2026-08-15

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

This change supplies reproducible simulated-source evidence only. It is not a
cold-filesystem or ambient-network measurement and does not accept a release
latency, tail, allocation, RSS, physical-device-I/O, or native DOC/XLS/PPT
performance claim. A configured service floor is a model accounting quantity,
not observed wall-clock service.

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

Release distributions and any paired latency result remain a separate evidence
commit.
