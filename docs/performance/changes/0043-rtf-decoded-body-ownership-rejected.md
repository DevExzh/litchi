# RTF decoded-body ownership handoff is rejected

Date: 2026-08-11

Production base: `c6e5ff7df1a4b1c0e8fb4b6c0985eed5b93fe8c8`

Disposition: measured and fully reverted. RTF production code is byte-identical
to the production base. OLE2, OOXML, ODF, iWork and IWA production code are
unchanged by this record. The ODF GenericArray deprecation cleanup remains in
the earlier accepted change 0041.

## Hypothesis

`Parser::flush_text_buffer` decodes each ordinary body block, copies the result
into the parser arena, and `RtfDocument::parse_string` later calls
`Cow::into_owned` while detaching the final document. A legacy-code-page decode
already returns an owned `String`, so the old path performs two additional
copies around that temporary allocation. Retaining the decoder allocation in
the final block should remove real work without changing text, byte offsets,
revision ranges or public ownership.

The fixed large CP-1252 corpus has 10,000 paragraphs, 529,999 decoded UTF-8
body bytes and one literal `0xe9` byte per paragraph. It is 560,063 source
bytes with SHA-256
`7157437b91b57aa50fd9725861c350f83eca84d0114341bcbc7bcda554fed50e`.

## Prototypes

The broad prototype retained every ordinary decoded body block directly.
Tracked insertions and deletions kept the arena path because revision records
borrow the same text while their ranges are assembled. It preserved the
existing effective-code-page selection, lossy body decoding, UTF-8 byte
offsets, overflow error, paragraph-content state and buffer-clear ordering.

That prototype removed the expected CP-1252 work. Heaptrack over two warmups
and 20 large opens reports 1,192,140 -> 951,876 allocation calls (-20.15%) and
479,524 -> 240,052 temporary allocations (-49.94%). Peak heap moves 22.95 ->
21.92 MiB (-4.49%). A 699/684-sample `perf record` comparison drops
`flush_text_buffer` from 2.19% exclusive share to below the 0.5% report cutoff;
no samples were lost. Uninstrumented maximum RSS averages 30,784 -> 30,912 KiB
(+0.42%, flat).

The primary plus independent confirmation used two A-B-B-A cycles, first with
50 warmups/500 samples per leg and then with 100 warmups/2,000 samples per leg.
Pooling 5,000 samples per state gives:

| Large CP-1252 open | Before | Broad prototype | Delta |
|---|---:|---:|---:|
| p50 | 3.225 ms | 3.126 ms | -3.08% |
| mean | 3.290 ms | 3.182 ms | -3.28% |
| p95 | 3.740 ms | 3.548 ms | -5.13% |
| p99 | 4.225 ms | 4.043 ms | -4.31% |

All four matched median pairs improved. A deterministic 10,000-resample
bootstrap puts the mean delta at [-3.59%, -2.97%] and the median delta at
[-3.35%, -2.82%]. This is real input-specific work elimination, but it was not
safe to accept without ordinary transport guards.

## Rejection guards

The same broad ownership handoff moves the final allocation of borrowed ASCII
from the compact post-parse ownership pass into 10,000 interleaved parser
iterations. A 20-warmup/200-sample-per-leg A-B-B-A guard rejected that change:

| Open guard | p50 delta | Mean delta | p95 delta |
|---|---:|---:|---:|
| Plain medium | +8.32% | +3.62% | -7.29% |
| Plain large | **+25.53%** | **+22.45%** | **+14.27%** |
| LZFu medium | +6.91% | +4.27% | -9.89% |
| LZFu large | +3.17% | +2.30% | -5.11% |
| CP-1252 medium | +3.65% | +3.55% | +9.19% |

Prepared operations that do not execute the parser also exposed binary-layout
instability: large CP-1252 full-text regressed 13.06% p50/14.13% mean and its
medium counterpart regressed 9.68%/12.52%. Exact stream-save and no-op segments
were mixed rather than consistently improved.

Two narrower prototypes retained only decoder-owned strings and left borrowed
ASCII arena-backed. A double discriminant inspection produced a statistically
real but immaterial 1.41% p50/1.11% mean CP-1252 improvement (4,000
samples/state); p99 regressed 4.33%. Replacing it with one consuming match
changed the same workload to +1.02% p50/+0.99% mean, with a bootstrap mean
delta interval of [+0.67%, +1.31%]. That compiler/code-layout sensitivity does
not meet the preregistered 3% materiality floor.

## Correctness and preservation

Every prototype retained non-body destination discard behavior, table parsing,
font-specific code pages, `\u` fallback handling, deletion invisibility,
insertion ranges, paragraph state, typed limits, exact raw/compressed no-op
bytes and final model ownership. The complete all-feature/all-target RTF suite,
focused revision and parse-limit suites, warning-denied Clippy/rustdoc and the
RTF fuzz-target build passed during the experiment.

A retained regression now proves that an incomplete Shift-JIS tail remains a
lossy U+FFFD body character while immutable `Document::to_bytes` returns the
exact original transport. This is correctness coverage only; it does not keep
any ownership prototype.

## Decision

Rejected and fully reverted. `flush_text_buffer` and the final model ownership
pass are byte-identical to the production base. No public API, dependency,
cache, runtime, lock, unsafe code, resource limit or semantic behavior changed.
The post-revert release build reproduced the frozen before executable exactly:
SHA-256 `671c9dcab035382fe23c8383a6c2cc50019d8204842069cac9ce6b0cfbb4335f`.
The raw JSON reports, digests and profile/memory command summaries are under
`docs/performance/results/` with the `rtf-decoded-body-` prefix.

Do not revive the broad handoff: its CP-1252 allocation win is outweighed by
the ordinary ASCII regression. A future attempt needs a materially different
ownership boundary that preserves the compact final-allocation pass, or a
larger parser redesign that removes enough work to dominate code-layout noise.

## Next non-iWork work

1. Attribute a distinct RTF parser frame rather than revisiting decoded block
   ownership.
2. Continue OLE2 final-owner attribution and source-backed OOXML publication
   without weakening independent readback.
3. Continue broader ODF source-backed reads and structural/resource-adding
   publication controls.
4. Keep iWork/IWA deferred while the `iwa-*` crates are modified elsewhere.
