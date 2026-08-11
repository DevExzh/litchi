# RTF ordinary-text byte delimiter scanning

Date: 2026-08-11

Production base: `af24e047c55bc16058d82df1c8c552b8bcf5a566`

Scope: private native RTF lexer ordinary-text scanning only. OLE2, OOXML,
ODF, iWork and IWA production code are unchanged.

## Disposition

Accepted. The lexer now finds ordinary-text delimiters with one byte pass
instead of decoding and advancing over every UTF-8 scalar twice. This changes
no public API, dependency, transaction, writer, limit, runtime, cache, lock or
unsafe-code boundary.

The fixed large plain corpus opens **17.23% faster at p50** and its complete
one-paragraph edit/save improves **14.65% at p50**. Plain, raw CP-1252 and
LZFu opens improve at both medium and large shapes. Allocation calls,
temporary allocations and peak heap are effectively flat.

## Measured bottleneck and work removed

The accepted transport-batching follow-up left
`Lexer::tokenize_with_spans` as a dominant RTF-owned large-open frame. The
matched before profile attributes 17.36% exclusive cycles to that frame. Its
ordinary-text loop called `current_char`, then `advance`, which called
`current_char` again. A long text token therefore rebuilt and validated the
remaining UTF-8 slice and decoded every scalar twice even though RTF syntax is
delimited by the five ASCII bytes `\\`, `{`, `}`, CR and LF.

The accepted loop scans the source bytes once and returns the first delimiter
with its byte offset. Structural delimiters remain for the next token. Physical
line breaks retain the established behavior: a line break after text is
consumed while the preceding borrowed text is returned; leading consecutive
line breaks are skipped; a trailing physical line break still produces the
established empty text token.

ASCII delimiter bytes cannot occur inside a UTF-8 continuation sequence. The
scan begins at a checked UTF-8 boundary and can therefore end only at another
valid boundary or EOF. Borrowed source ranges, exact token spans and the typed
invalid-private-cursor error remain intact. No extra dependency or unsafe code
is used.

## Primary matched latency

Environment: release profile, Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator and CPU 2 pinned with `taskset`. The fixed
large plain corpus contains 10,000 paragraphs and 540,051 source bytes; its
SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.

The before executable SHA-256 is
`4f906af68f226924ca549162cece4a8bac9d9e87885af8808faa65c9d3ab4c5a`;
the final after executable SHA-256 is
`ba1b6041e7ca0ebf721700c25bad90884cb6bf2ada665989e004b4d0bf7168d9`.

The primary run used 50 warmups and 500 samples per leg in
before-A, after-A, after-B, before-B order. Pooling both legs yields 1,000 raw
samples per state.

| Large plain RTF workload | Before p50 | After p50 | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|---:|---:|
| Open | 2.479 ms | 2.052 ms | **-17.23%** | **-17.99%** | **-18.51%** | -20.79% |
| One-paragraph edit/save | 7.554 ms | 6.447 ms | **-14.65%** | **-14.84%** | **-16.34%** | -15.28% |

Raw reports:

- [`before A`](../results/abba-rtf-byte-delimiter-final-before-a.json),
  SHA-256 `8066d06131b250737a5ef8f7e9a6bdb912792666a14c546bc37d143d448c38d0`;
- [`after A`](../results/abba-rtf-byte-delimiter-final-after-a.json),
  SHA-256 `1d87c09d2606cfd89f12a6574850afade4b98654aaffbd297218dec09148df0a`;
- [`after B`](../results/abba-rtf-byte-delimiter-final-after-b.json),
  SHA-256 `6013a9525583761c1339482e7cdbd3727bab43c9ac1afe0701edb921eccbdcab`;
- [`before B`](../results/abba-rtf-byte-delimiter-final-before-b.json),
  SHA-256 `35a5798425c34d89c8a5184da0e114747f68e36e538473a0d224ea039e5b4c1c`.

## Variant and regression guards

A separate four-leg run used 20 warmups and 100 samples per leg across medium
and large plain, raw CP-1252 and LZFu corpora. All six open cells improve:

| Open guard | Medium p50 | Large p50 |
|---|---:|---:|
| Plain | -21.06% | -23.64% |
| Raw CP-1252 | -12.89% | -11.06% |
| LZFu | -23.93% | -19.39% |

Plain one-edit/save improves 24.50% at medium and 14.03% at large. Large
plain and raw CP-1252 full text improve 7.70% and 7.41%. A dedicated
2,000-sample/state LZFu follow-up resolves mixed-run full-text noise to
**-6.51% p50 / -6.02% mean / -5.02% p95**.

The same dedicated LZFu run discloses one narrow exception: the prepared exact
no-op edit/save segment moves from 4.526 to 4.816 us (**+6.41% p50**, +7.29%
mean). The timer deliberately begins after decompression and parsing, so the
changed lexer is not executed inside that measured segment; the movement is a
0.290 us cache/layout after-effect. The complete large LZFu open, which is the
changed portion of an open-then-no-op workflow, improves 19.39% at p50. The
exception is accepted because the end-to-end workflow is materially better,
the timed no-op implementation and exact output are unchanged, and memory
remains flat. It is not presented as a no-op improvement.

Raw variant guards are the four
[`final-guards`](../results/abba-rtf-byte-delimiter-final-guards-before-a.json)
ABBA reports and the four dedicated
[`LZFu guards`](../results/abba-rtf-byte-delimiter-final-lzfu-guard-before-a.json).
The final one-sample coverage report contains 25 unique tiny identities across
plain, raw CP-1252, LZFu and the producer watermark:
[`variant smoke`](../results/rtf-byte-delimiter-variant-smoke.json), SHA-256
`159013de5fd865a80708b1681d5f68df4788e6b753959ba76829d444d8fd65b6`.

## Profile, counters and memory

Matched large-open `perf record` runs used ten warmups and 120 samples.
`Lexer::tokenize_with_spans` falls from 17.36% to 11.06% exclusive cycles
(36.29% lower share). The former scalar `current_char`/`advance` path no longer
accounts for ordinary-text traversal.

Whole-process `perf stat` ABBA runs used 20 warmups and 200 large open plus
one-edit/save samples per leg:

| Counter, A+B | Before | After | Delta |
|---|---:|---:|---:|
| Task clock | 10,265.2 ms | 9,161.59 ms | -10.75% |
| Cycles | 50,165,661,600 | 44,678,316,517 | -10.94% |
| Instructions | 168,469,495,644 | 132,635,245,533 | **-21.27%** |
| Branches | 44,481,882,598 | 31,453,907,483 | **-29.29%** |
| Branch misses | 22,425,328 | 22,171,693 | -1.13% |
| Cache references | 3,748,228,630 | 3,722,945,075 | -0.67% |
| Cache misses | 196,593,679 | 210,158,603 | +6.90% |

The cache-miss increase is disclosed; it does not produce a latency, heap or
RSS regression in the changed workflows, while instructions and branches fall
materially.

Heaptrack over two warmups and 20 large samples reports 2,308,467/2,308,468
allocation calls, 460,118/460,119 temporary allocations, identical 56.98 MiB
peak heap and 544 leaked bytes. The one-call difference over the complete
instrumented process is immaterial. Heaptrack RSS is 60.14/60.45 MiB (+0.52%).
Uninstrumented GNU Time ABBA maximum RSS is 56,300/56,428 KiB (+0.23%, flat),
while user time falls from 3.79/3.60 s to 3.14/3.15 s.

Raw evidence is in
`rtf-byte-delimiter-{before,final-after}-perf-report.txt`,
`rtf-byte-delimiter-final-perf-stat-*.csv`,
`rtf-byte-delimiter-{before,final-after}-heaptrack.txt`, and
`rtf-byte-delimiter-{time-before,final-time-after}-*.txt` under `results/`.

## Preservation, safety and verification

The change retains code-page decoding, font/code-page controls, Unicode
controls and fallbacks, hexadecimal escapes, binary payloads, unknown/opaque
syntax, token/source/decompression limits, immutable source identity, exact
no-op bytes, checked edits, durable patch/inverse and stale-source behavior,
candidate parse/readback, LZFu publication refusal and forward-only sink
semantics.

Focused lexer coverage pins multibyte UTF-8 text, every structural delimiter,
CR/LF/consecutive physical line breaks and exact token spans. Existing tests
retain escaped delimiters, control-word boundaries, invalid UTF-8 cursor
errors, raw CP-1252 behavior, token limits and trailing line breaks.

Final gates:

- the complete `litchi-rtf --all-features` suite passes, including 296 library
  unit tests, every integration suite and nine doctests;
- warning-denied all-target/all-feature Clippy and warning-denied crate rustdoc
  pass;
- the `parse_rtf` fuzz target and its production dependency graph compile
  offline;
- the unchanged benchmark harness passes all 29 tests and warning-denied
  all-target Clippy;
- the 25-row final transport/producer smoke passes with exact corpus hashes;
- exact-file formatting and repository diff hygiene pass.

Existing CI already exercises the 25-row tiny RTF matrix on pushes and the
44-row tiny-plus-large matrix on scheduled/manual release runs, so no CI case
or threshold was weakened or renamed.

## Next bounded work

1. Attribute the deferred ODT compact-audit archive copies before changing
   them.
2. Add a source-backed PPTX owning editor and media-rich control before a
   one-slide overlay publisher; do not attach consuming publication to the
   cloneable read facade.
3. Add a bounded multi-Part OPC publisher before source-backed XLSX cell
   edits, retaining workbook recalculation and calculation-chain policy.
4. Add DOC owner-stage attribution before another OLE2 prototype; the common
   layer has no remaining accepted handoff candidate.

iWork/IWA remains explicitly deferred while other agents modify `iwa-*`.
