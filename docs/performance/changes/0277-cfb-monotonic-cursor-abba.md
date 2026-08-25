# Change 0277: CFB monotonic stream cursor and XLS ABBA evidence

Date: 2026-08-25

Status: production optimization retained as a measured reusable enabler;
`performance_claim: none`

## Production change

`litchi-cfb` now owns a non-`Clone`, forward-only
`SharedOleStreamCursor<'_>` for validated FAT and MiniFAT streams. The cursor
retains private chain position, translates logical offsets once, batches only
physically contiguous sectors, and supports exact reads and forward skips. It
does not materialize the MiniStream cache. The final partial MiniFAT sector is
capped to the remaining root-stream bytes, including the valid 4,095-byte
root-tail case.

Construction and skip traverse only the immutable, already validated CFB
catalog and publish no source bytes. Every exact payload read retains the
existing before/after source-version fence. A mutation observed after an I/O
error returns `SourceChanged` in preference to the payload error; unchanged
sources preserve the original typed error. Cursor state commits only after a
successful final fence, and callers must discard a partially filled output
buffer after any error.

The private XLS source owner uses the cursor for Workbook-global header
preflight and selected-worksheet frame traversal. Four-byte BIFF headers use a
stack buffer, supported payloads are read exactly, and unknown payloads are
skipped without source I/O. The final global span, FILEPASS/EOF behavior,
limits, cancellation boundaries, STRING/CONTINUE ordering, duplicate-last
cells, and selected-sheet-only locality are unchanged. No CFB cursor or
physical identifier crosses an ordinary XLS or facade CRUD signature.

The production revisions are:

- `634803f39`: add the monotonic shared stream cursor and XLS integration
- `eac9b7518`: remove construction-only source probes while retaining every
  payload-read fence

## Correctness and architecture gates

Independent validation passed:

- 282 `litchi-cfb` tests: 263 unit, 13 sequential-writer, and 6 storage-move
- 29 all-feature `litchi-xls` source-backed integration tests
- all-feature CFB and XLS library/test checks
- strict CFB Clippy and targeted XLS library/source-backed Clippy
- strict CFB and XLS rustdoc
- the crate-boundary gate for 64 packages and 240 declarations

Focused regressions cover FAT and MiniFAT cursor reads, fragmented and
contiguous chains, forward skips, error atomicity, mutation before/during/after
reads, error-versus-freshness precedence, valid partial MiniStream tails, and
an unknown 4,096-byte BIFF payload that is skipped without a payload read.
Architecture, cursor-fence, XLS frame/locality, and final diff reviews found no
P0-P2 blocker. Full XLS Clippy still encounters an unrelated pre-existing
`let_and_return` warning in `tests/validation.rs`; the changed XLS paths are
clean.

## Strict A1/B1/B2/A2 run

The control is `ebaaf057ad7773da7d9e061594c4c5100ff327ae`, release binary
SHA-256 `e046fa377f6e3e4f712a124d454e434e7da2b7c0af72ccd02aefa84673b22017`
and size 51,051,864 bytes. The candidate is
`eac9b7518f58b1796926adeaa26a5b9627ccf873`, release binary SHA-256
`9a9f6cd33e0c4368920f470a41765f007f22b637de716b0306613325e4affd4d`
and size 51,072,448 bytes. Both worktrees were clean. The canonical
configuration SHA-256 is
`a187ebf7988abcf5050f55498f0a6adca400d76c708012c04f89b8bbe0080109`.

All legs used Rust 1.95 release builds, CPU 2, one worker, fresh isolated warm
children, 20 warmups, 500 retained samples, and one six-selector dispatch.
Positive percentages below mean that the candidate is faster. Drift ceilings
were 5% for p50 and mean, 10% for p95, and 15% for p99.

The corpus is the 16,995,840-byte opaque-heavy XLS archive, SHA-256
`6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53`.
Its 80,946-byte Workbook stream has SHA-256
`c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041`.

| Source-backed lifecycle | A1 -> B1 p50 / mean / p95 / p99 | A2 -> B2 p50 / mean / p95 / p99 |
|---|---:|---:|
| open | +1.52% / +1.47% / +1.68% / +2.23% | +1.54% / +1.46% / +0.77% / -1.25% |
| open + list | +1.54% / +2.03% / +2.49% / +7.18% | +1.00% / +0.36% / -4.65% / -5.31% |
| open + one cell | +2.01% / +0.85% / -1.81% / -31.64% | +1.63% / +1.59% / +1.25% / +1.00% |

The source p50 and mean cells pass both paired-direction stability gates for
all three lifecycles. Open p95 also passes. List and one-cell p95/p99 tails do
not pass, and no tail claim is made. In particular, one-cell A1/B1 p99 is
31.64% slower and list A2/B2 p99 is 5.31% slower. Both exceed the program's
5% review trigger; the reverse legs disagree rather than reproduce those
regressions, so they remain explicit follow-up evidence instead of being
averaged away.

The eager controls are environment guards rather than attributable cursor
results:

| Eager lifecycle | A1 -> B1 p50 / mean / p95 / p99 | A2 -> B2 p50 / mean / p95 / p99 |
|---|---:|---:|
| open | +2.27% / +2.33% / +3.19% / -0.99% | -0.52% / -0.86% / -5.93% / -4.03% |
| open + list | +1.48% / +4.17% / +17.46% / +27.15% | -0.80% / -1.27% / -0.62% / -6.91% |
| open + one cell | +6.32% / +7.14% / +13.78% / +16.09% | +1.77% / +1.95% / +1.85% / +0.86% |

Across 24 selector/statistic comparisons, 10 pass both directions, 14 are
rejected, and one rejected eager-control cell is adverse in both directions.
Eager cells cannot establish cursor causality.

## Exact logical-work result and decision

Every sample in every A/B/B/A leg records the same source work for control and
candidate:

| Source selector | Calls / bytes | CFB structural | Workbook globals | Selected worksheet | Version checks |
|---|---:|---:|---:|---:|---:|
| open | 334 / 138,459 | 265 / 136,704 | 69 / 1,755 | 0 / 0 | 168 |
| open + list | 334 / 138,459 | 265 / 136,704 | 69 / 1,755 | 0 / 0 | 162 |
| open + one cell | 362 / 138,593 | 265 / 136,704 | 69 / 1,755 | 28 / 134 | 224 |

Unselected-worksheet and opaque-payload calls/bytes remain zero. The mechanism
therefore removes repeated in-memory chain walking and small header allocation;
it does not reduce logical reads, bytes, or source-version observations. The
mutex-instrumented source path remains roughly 11-12x eager, and its range
tracking can dominate the small 0.36%-2.03% central gains.

The cursor is retained because the stable central cells and coherent exact
mechanism make it a reusable low-level enabler, not because this run proves a
selector-wide improvement. `performance_claim: none`: the program comparator
currently rejects this lifecycle corpus as a native-numeric selector mismatch,
and the rejected tails preclude a registered selector claim. No FileSource,
physical-I/O, cold-cache, allocation/RSS, peak-memory, producer, encrypted,
edit/save, broad XLS, or broad CFB claim follows.

The raw reports, corrected projection, and hash/size manifest are retained in
[`results/0277-cfb-monotonic-cursor-abba-20260825/`](../results/0277-cfb-monotonic-cursor-abba-20260825/).
The diagnostic one-sample provenance probe is excluded. The next evidence
batch should separate owned, atomic-only, tracked, FileSource, facade, CFB-open,
global-parse, and selected-sheet phases before considering bounded same-stream
span batching.
