# Change 0147: CFB MiniFAT `open_stream` release ABBA

Date: 2026-08-16

Status: scoped one-shot simulator result accepted; repeated-work tradeoff
retained; no generic local-latency or resource claim.

## Compared revisions and matrix

This record measures the 12 selectors introduced by
[change 0146](0146-cfb-open-stream-evidence.md) against the production state
immediately before `3375729f4` while keeping the harness source identical.
The clean control is temporary measurement commit `230cc51a1`, made from
candidate/harness commit `24dfc0a38` by restoring only
`crates/litchi-cfb/src/shared.rs` to `88eda0fa3`. The shared harness source has
SHA-256 `7e77315dced0be72e8b3f65f0c6825a9c24f9a7e1bd8cd9bba861a5e010b4cb8`;
the control and candidate `shared.rs` hashes are
`db9855efa987a4ab858fa416c0fdebf5099131ec0baf1b9aebc6874134f047a1`
and `ae74089e4abdf672b9ec1c3010db007cbc9ba9a3736de1ac66e21b62781661b9`.

The release binaries were built with Rust 1.95.0 from clean detached
worktrees. The control is 40,174,528 bytes with SHA-256
`d5ba999bfbd876df667e109db47a50fd5448f3f5b54fa807c43804e6928de128`;
the candidate is 40,183,272 bytes with SHA-256
`404f2852bad7c3670bcff063222be2854acec1246c816ab5620aee79b2c54eac`.
Four fresh CPU-2 processes ran in strict
`A1 control, B1 candidate, B2 candidate, A2 control` order. Every process used
20 warmups and 200 samples for each of 24 records: 12 selectors across the
`many-small` and `wide-root` shapes. The retained matrix therefore contains
19,200 samples. All reports record clean worktrees, affinity `2`, exact corpus
and output hashes, stable source versions, and the exact
`OleError::StreamNotFound` refusal.

The configured simulator uses 100 us fixed latency, 25 us request overhead,
50 MiB/s bandwidth, and a 4 KiB maximum physical request. It is a deterministic
harness model, not a network, device, syscall, or physical-I/O observation.

## Exact one-shot source work

All 19,200 raw event sequences satisfy their declared formulas. For the
one-shot cells, the parent reads the complete root Mini Stream while the
candidate reads one exact target range:

| Shape / target | Control returned bytes | Candidate returned bytes | Candidate range `[start,end)` |
|---|---:|---:|---:|
| many-small / 36 | 261,184 | 36 | `[261632,261668)` |
| many-small / 4,095 | 265,216 | 4,095 | `[261632,265727)` |
| wide-root / 36 | 2,096,192 | 36 | `[2096640,2096676)` |
| wide-root / 4,095 | 2,100,224 | 4,095 | `[2096640,2100735)` |

Every request returns the exact expected payload length and SHA-256. These are
logical positional-source events, not physical storage reads.

## Accepted configured-simulator result

The table reports candidate improvement percentages in the two adjacent ABBA
directions. Positive values are faster. `Operation p50` covers only the public
`open_stream` call; `total` is the checked fresh parser-open plus operation
sum.

| Shape / target | Operation p50, A1->B1 / B2->A2 | Total p50 | Total p95 | Total p99 | Total mean |
|---|---:|---:|---:|---:|---:|
| many-small / 36 | 98.900% / 98.896% | 62.704% / 62.676% | 62.534% / 62.393% | 62.214% / 62.321% | 62.688% / 62.653% |
| many-small / 4,095 | 98.449% / 98.448% | 62.274% / 62.284% | 62.320% / 62.197% | 62.866% / 64.324% | 62.329% / 62.316% |
| wide-root / 36 | 99.858% / 99.858% | 63.980% / 64.028% | 63.959% / 63.917% | 63.660% / 63.827% | 63.965% / 64.033% |
| wide-root / 4,095 | 99.800% / 99.799% | 64.039% / 63.951% | 63.821% / 63.845% | 63.846% / 63.927% | 64.005% / 63.942% |

Both directions agree at p50, p95, p99, and mean for every named one-shot
cell. The claim is limited to this simulator configuration, corpus, build,
host, and operation. The local in-memory positional source also shows large
one-shot operation-interval reductions, but several small-duration tail and
control distributions drift materially; no generic local wall-clock claim is
accepted from those samples.

## Repeated-work tradeoff

The candidate deliberately consumes one direct opportunity before the second
call materializes the root cache. Exact repeated source work is therefore
`[L,R,0...]` rather than the control's `[R,0...]`. Under the configured model,
that extra request is visible on the many-small corpus:

| Target / operation | Total p50, A1->B1 / B2->A2 | Total mean | Total p99 |
|---|---:|---:|---:|
| 36 / repeat-3 | -0.534% / -0.903% | -0.446% / -1.040% | -0.157% / -2.618% |
| 36 / repeat-8 | -0.911% / -0.384% | -1.221% / -0.295% | -9.525% / +2.709% |
| 4,095 / repeat-3 | -0.790% / -0.948% | -0.867% / -1.064% | -2.833% / -2.085% |
| 4,095 / repeat-8 | -0.624% / -0.809% | -0.733% / -0.760% | -1.947% / -0.062% |

Negative values are regressions. The one 9.525% p99 regression is an explicit
review trigger, but it reverses in the paired direction and coincides with
same-implementation tail drift, so it is not accepted as a material regression.
The smaller p50/mean regressions are
consistent with the exact extra request and are retained rather than averaged
away. Wide-root repeated cells are near neutral under the configured model.
No generic repeated-read improvement is accepted. A follow-up should evaluate
whether exact same-target repeats can stay direct while different-stream,
bulk, and concurrent callers retain the bounded root-cache takeover.

## Artifacts and claim boundary

The [compact summary](../results/cfb-open-stream-abba-0147-summary.json) has
SHA-256 `95b118a04ca31161ff26ae76d3eb7c8c5783b9e0438bd01c99b41b5f76c62e80`
and contains all phase percentiles, per-invocation
statistics, source-work vectors, simulator floors, environment, revisions,
binary identities, and claim decisions. Complete raw reports are retained as:

- [A1 control](../results/cfb-open-stream-a1-control-0147.json.zst),
  SHA-256 `e034f09b8907443f06ffcba8edc9ede651f53260d86c4d5151128372043fcb5b`,
  598,660,723 bytes after decompression;
- [B1 candidate](../results/cfb-open-stream-b1-candidate-0147.json.zst),
  SHA-256 `1a75ca44ed2f146975364361d1ca1cf82f3a8110bc13b53e27caba1a38fa8de2`,
  467,547,874 bytes after decompression;
- [B2 candidate](../results/cfb-open-stream-b2-candidate-0147.json.zst),
  SHA-256 `3ca54066fa7b347c91dca3f0a30245e70a60cbffb789b0429a0bca61e7360ae7`,
  467,547,662 bytes after decompression;
- [A2 control](../results/cfb-open-stream-a2-control-0147.json.zst),
  SHA-256 `dd559e6ab2d327aae0fe8611a79835bff42844397d0cb6f03d68bf312afcdc59`,
  598,661,070 bytes after decompression.

No allocation, RSS, peak-memory, physical-I/O, cold-cache, remote/network,
device, decompression, DOC/XLS/PPT semantic, OOXML, ODF, RTF, or iWork result
is claimed. Bulk, concurrency, direct-failure/retry, ineligible-root, and FAT
`open_stream` matrices remain follow-up work.
