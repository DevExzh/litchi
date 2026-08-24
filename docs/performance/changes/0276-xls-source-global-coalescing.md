# Change 0276: XLS source-global coalescing

Date: 2026-08-25

Status: production optimization retained from clean decision smoke;
`performance_claim: none`

## Production change

The BIFF8 source-backed XLS owner now preflights Workbook globals with
header-only reads, validating exact frame bounds, configured byte/record
limits, FILEPASS, and a non-empty EOF before reading any payload. Once the
first global EOF is known, it performs one exact `[0, global_end)` read and
feeds the existing semantic parser from that bounded buffer. On the fixed
opaque-heavy corpus this changes global logical ranges from `2G = 136` to
`G + 1 = 69` while preserving the same 1,755 logical global bytes. It never
prefetches a worksheet, encrypted payload, malformed tail, or opaque stream.
The new allocations are fallible and remain under existing finite limits.

The facade now constructs one `SharedOleFile` catalog for OLE host
classification, hands that same immutable catalog to an explicit
`litchi_xls::raw` low-level constructor, and retains it for metadata.
Ordinary CRUD signatures remain CFB-free. The raw handoff derives the exact
source `Arc` from the catalog, rechecks source freshness and the configured
input ceiling, and cannot pair a catalog with an unrelated source having a
coincidentally equal version token. A counting regression proves that
path detection, XLS globals, and later metadata reuse one catalog; OLE host
precedence remains Word, then PowerPoint (including `Current User`), then
Excel.

The supporting CFB range reader also now runs its documented post-read source
fence on error paths. A concurrent mutation returns `SourceChanged` in
preference to a payload error; an unchanged source preserves the original
typed error and can be retried.

## Correctness and locality gates

Focused tests cover exact global-span coalescing, zero worksheet overlap,
large/truncated FILEPASS without payload reads, truncated headers without
overread, stale raw handoff, tighter input limits, source `Arc` identity,
single-catalog facade reuse, and path-level host precedence with all three
legacy owners enabled. Existing duplicate-last cell semantics, SST/CONTINUE,
late codepage, worksheet EOF traversal, cancellation, record/scan limits,
materialization, and selected-sheet-only locality remain covered.

Verification passed:

- 256 `litchi-cfb` library tests and its all-feature lib/test check
- 28 `litchi-xls` source-backed integration tests and its all-feature check
- 13 xls-only and 13 `doc+ppt+xls` facade source-path tests
- facade lib/test checks for `xls+xlsx`, `xls+ods`, `xls+docx`, and
  `xls+pptx+xlsb`
- two independent final reviews with no P0-P2 findings

## Clean matched decision smoke

Control revision
`c60c67d39ead9e40f5c02067c5f91ebbab8b099c` used release binary SHA-256
`104e863a2467ecb7484f78f2f19bca00c45a88c5db1ae838bf991da6ace8624a`.
Candidate revision
`ebaaf057ad7773da7d9e061594c4c5100ff327ae` used release binary SHA-256
`e046fa377f6e3e4f712a124d454e434e7da2b7c0af72ccd02aefa84673b22017`.
Both detached worktrees were clean. CPU-2 A1/B1/B2/A2 legs used one worker,
three warmups, 30 fresh isolated warm samples per selector, and a single
six-selector dispatch so every eager/source pair used the same corpus.

The corpus is the 16,995,840-byte
`litchi-xls-comments-opaque-heavy-v1` archive, SHA-256
`6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53`.
Its 80,946-byte Workbook stream has SHA-256
`c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041`.

| Source-backed lifecycle | A1 -> B1 p50 / mean / p95 / p99 | A2 -> B2 p50 / mean / p95 / p99 |
|---|---:|---:|
| open | -13.44% / -12.97% / -11.17% / -3.48% | -14.51% / -14.41% / -14.64% / -14.88% |
| open + list | -13.59% / -13.56% / -14.27% / -14.48% | -13.88% / -13.85% / -14.06% / -13.63% |
| open + one cell | -12.87% / -12.59% / -13.03% / -12.97% | -12.08% / -12.18% / -11.63% / -13.65% |

All 24 source-backed statistic cells improved. The paired eager controls had
no adverse movement in both directions: open p50 moved -4.33% / -0.90%, list
+0.75% / -5.32%, and one-cell -0.83% / -2.06%. Deterministic logical evidence
per sample changed as follows:

- open/list total calls: `401 -> 334` (-16.71%)
- one-cell total calls: `429 -> 362`, retaining 28 selected-sheet calls
- global calls: `136 -> 69` (-49.26%) at equal 1,755 global bytes
- source-version checks: open `299 -> 168`, list `293 -> 162`, one-cell
  `355 -> 224`
- total logical bytes remain 138,459 for open/list and 138,593 for one-cell
- unselected worksheet and opaque-payload calls/bytes remain zero

The candidate is still approximately 11x eager for open/list and 12x eager
for one-cell in this mutex-instrumented logical-range harness. This retained
30-sample package is sufficient for the batch keep/revert decision, but not
for a registered performance claim: it does not meet the program's
500-sample strict acceptance protocol, and it does not measure uninstrumented
`FileSource`, physical I/O, cold cache, allocation/RSS, or peak memory.
No broad XLS, CFB, producer, encrypted-file, edit/save, or all-BIFF-version
claim follows.

The four combined reports and their identity manifest are retained under
[`results/0276-xls-source-global-coalescing-20260825/`](../results/0276-xls-source-global-coalescing-20260825/).
The next hotspot is a bounded monotonic CFB stream session or validated
same-stream span batching to reduce repeated FAT-chain walks, version probes,
and selected-worksheet header/payload requests without weakening source,
limit, malformed-tail, cancellation, or duplicate-last behavior.
