# Change 0125: MiniFAT physical-run evidence controls

Date: 2026-08-15

## Scope

The standalone performance harness adds two opt-in selectors for a distinct
4095-byte MiniFAT target:

- `cfb_selective_mini_4095_legacy_read`
- `cfb_selective_mini_4095_shared_read`

They use the same deterministic `many-small` and `wide-root` sibling shapes as
the existing 36-byte MiniFAT control and retain that control unchanged. The
4095-byte target is the largest stream below the CFB MiniFAT cutoff and spans
64 logical 64-byte mini-sectors (eight regular 512-byte sectors). Its final
position makes the difference
between full root-mini-stream materialization and bounded positional reads
visible without adding a semantic DOC, XLS, or PPT dependency.

The legacy selector uses a cursor and materializes the complete target stream.
The paired selector opens the same archive through `SharedOleFile`, allocates
an exact 4095-byte caller buffer, and uses `read_stream_range`. The resulting
source range vector records whether contiguous physical root-sector runs were
coalesced. The manifest retains the archive SHA-256, target SHA-256, target
length, entry count, logical payload bytes, generator, shape, and target path.

## Recorded evidence

Every sample records separate `open_ns`, `read_ns`, and `total_ns` values. Open
and selected-read source counters are reset between stages. Each stage records
read calls, returned bytes, sorted range sizes, returned payload length, and the
selected payload SHA-256. The two implementations are measured against the
same generated archive; corpus construction, validation, and hash generation
are outside the timed stages. The report does not treat these counters as a
private cache/materialization metric.

The focused correctness gate requires the 4095-byte payload and hash to match
the manifest, legacy source bytes to exceed the selected payload, positional
source bytes to equal exactly 4095, and the positional source range vector to
be one exact 4095-byte request. This exposes request amplification while
remaining independent of a particular latency result.

These selectors are evidence only. No speedup, tail-latency, physical-I/O,
allocation, peak-RSS, cold-cache, high-latency-source, or semantic native
Office claim is accepted from the focused smoke. Release ABBA and resource
attribution remain required before any performance conclusion.

## Gates and exclusions

The default matrix remains 36 cases / 198 result records. The two new names
raise the selectable case-name count from 243 to 245. The six CFB selective
selectors emit only the `many-small` and `wide-root` shapes. iWork remains
deferred.

Focused execution:

```sh
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml selective_cfb -- --nocapture
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 1 --shape many-small,wide-root \
  --case cfb_selective_mini_legacy_read,cfb_selective_mini_shared_read,\
cfb_selective_mini_4095_legacy_read,cfb_selective_mini_4095_shared_read,\
cfb_selective_fat_legacy_read,cfb_selective_fat_shared_read \
  --json target/perf/cfb-minifat-run-smoke.json
```

The smoke output is a reproducibility artifact and is not a release benchmark
result.
