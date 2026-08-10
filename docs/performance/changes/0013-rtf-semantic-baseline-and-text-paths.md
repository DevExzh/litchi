# Change 0013: RTF semantic baseline and text-path work elimination

Date: 2026-08-10

## Decision

Accept seven opt-in native RTF semantic benchmark cases and three private work
eliminations on their measured paths:

- `RtfDocument` records the total UTF-8 byte length while it already detaches
  parsed blocks. First full-text materialization then allocates exactly once
  and copies the retained blocks in one pass, instead of allocating a temporary
  `Vec<&str>`, scanning that vector for length, and then joining it.
- `RtfWriter::write_text` emits contiguous ordinary ASCII spans with one
  `write_all` call per span. Escapes, control bytes, Unicode fallback controls,
  delimiters, and sink errors retain their existing behavior.
- Text-only commits do not build two unused paragraph-alignment vectors or scan
  uniform bold state. A valid paragraph selector also returns as soon as its
  range is known instead of counting the rest of the story.

Full-text materialization remains lazy and cached by the ordinary immutable
`Document` facade. Property edits still execute all alignment and bold scans,
candidate validation and complete semantic readback remain mandatory, and
exact no-op commits still share the source snapshot.

An earlier candidate that pre-scanned the full block vector and then appended
directly was removed: pooled medium full-text p50 improved 16.95%, but the
10,000-paragraph p50 regressed 25.85%. Recording the byte length during the
existing detach pass removes that extra cache-unfriendly block scan and clears
the large guardrail.

## Benchmark coverage

The new cases use only `litchi_rtf::Document` and its public transaction and
streaming APIs:

1. owned-byte open;
2. lazy paragraph enumeration;
3. one middle paragraph and its text;
4. first complete-text materialization;
5. exact immutable stream save;
6. exact empty-edit commit and stream save; and
7. one checked middle-paragraph edit, commit and stream save.

The direct deterministic ASCII corpus contains 24, 200, or 10,000 paragraphs
and 1,347, 10,851, or 540,051 source bytes. Its stable SHA-256 values are
`ee4a5c5b...e62328`, `48b7dd8b...31999`, and `957645f9...6e02e`.
Every output is reopened and every paragraph plus complete text is verified.
Exact saves record 1,347/10,851/540,051 accepted bytes in one write; the large
changed save records 540,052 bytes in one write. The default 36-case / 198-row
matrix is unchanged; the harness now exposes 88 selectable cases.

## Matched latency result

The release executables were frozen on production base `1793c089f` after the
completed harness existed in both states:

- before SHA-256:
  `1d581ba92943905d852e112e070ff02a77b7dfca6b0d56cfe3145ab26b3385d3`
- after SHA-256:
  `793d01bc572a56ad88e5ba80bbe754621c557f6780d8529d1159d9c20d1d5f2f`

Both states ran pinned to CPU 2 in before-A, after-A, after-B, before-B order,
with 30 warmups and 250 measured samples per leg. The table pools both legs
(500 samples per state). Mean intervals are two-sided 95% intervals. The raw
reports mark the tree dirty because this batch and an unrelated pre-existing
documentation edit were uncommitted.

| Case | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| Full text, medium | 0.521 / 0.741 / 0.861 us | 0.321 / 0.350 / 0.541 us | **-38.39%** | 0.538 us (0.531-0.545) | 0.331 us (0.327-0.334) | **-38.61%** |
| Full text, large | 33.095 / 51.771 / 63.468 us | 24.134 / 40.848 / 47.997 us | **-27.08%** | 35.172 us (34.470-35.873) | 26.249 us (25.621-26.878) | **-25.37%** |
| One edit/save, medium | 195.970 / 227.463 / 264.826 us | 130.511 / 172.448 / 217.489 us | **-33.40%** | 200.202 us (198.766-201.639) | 137.946 us (136.369-139.523) | **-31.10%** |
| One edit/save, large | 12.408 / 13.511 / 14.767 ms | 9.208 / 10.180 / 11.164 ms | **-25.79%** | 12.499 ms (12.440-12.557) | 9.307 ms (9.266-9.349) | **-25.53%** |

The retained-length bookkeeping also runs during open. A separate 500-sample
pooled guard reports medium open p50 83.864 to 84.666 us (+0.96%) and large
open 3.832 to 3.963 ms (+3.41%); p95 moves -18.51% and +3.32%. No open
regression exceeds the 5% guard.

Raw targeted samples:
[`before A`](../results/abba-rtf-text-before-a.json),
[`after A`](../results/abba-rtf-text-after-a.json),
[`after B`](../results/abba-rtf-text-after-b.json), and
[`before B`](../results/abba-rtf-text-before-b.json).

Open guard samples:
[`before A`](../results/abba-rtf-open-before-a.json),
[`after A`](../results/abba-rtf-open-after-a.json),
[`after B`](../results/abba-rtf-open-after-b.json), and
[`before B`](../results/abba-rtf-open-before-b.json).

The complete seven-case matrix is retained as
[`before`](../results/rtf-semantic-before.json) and
[`after`](../results/rtf-semantic-after.json). Its 50-sample rows are useful
for broad smoke and sink counters; the higher-sample ABBA records above are
the decision evidence for sub-microsecond text work.

## Allocation and memory result

Heaptrack over 100 large one-edit/save samples reports allocation calls falling
from 6,118,986 to 6,118,279 (707 fewer) and temporary allocations from
1,020,608 to 1,020,406 (202 fewer). Peak heap is flat at 56.98 MB. These are
whole-process totals dominated by two complete RTF parses per measured edit;
the removed full-text fragment vectors and property vectors are therefore a
small share of the total allocation count.

Heaptrack's instrumented RSS varied from 55.44 to 65.86 MB and is not used as
an RSS claim. A reverse-order uninstrumented GNU Time guard reports 54,120-
54,376 KiB before and 54,552 KiB after; the maximum-to-maximum delta is +0.32%,
treated as flat noise and below the 5% regression threshold.

Profiler summaries:
[`before`](../results/heaptrack-rtf-before.txt),
[`after`](../results/heaptrack-rtf-after.txt), and
[`GNU Time RSS`](../results/time-rtf-rss.txt).

## Correctness and contract gates

- Fragmented formatting blocks retain exact concatenation and natural
  separators; ordinary ASCII is emitted in one chunk; every special escape
  spelling and delimiter still round-trips.
- Valid paragraph selection leaves the lazy paragraph-count cache empty.
  Existing scalar text, paragraph-property, durable patch, inverse, stale-base,
  opaque-preservation and failure-atomic transaction tests remain enabled.
- A forward-only sink that accepts a prefix and then fails returns the sink
  error without mutating the immutable source snapshot.
- CI watches `crates/litchi-rtf`, runs all seven tiny cases on pushes and pull
  requests, and publishes 14 tiny/large release records on scheduled or manual
  runs.
- The complete all-feature RTF suite passes (291 library unit tests, every
  integration suite and nine doctests), as do both library and all-target
  warning-denied Clippy. The standalone harness passes 22 tests, all-target
  warning-denied Clippy, formatting, and the seven-record release smoke.
- No public archive type, dependency inversion, unsafe code, ambient input,
  hidden runtime, global lock, or iWork/IWA change is introduced.

## Remaining limitations and next audits

The corpus is deterministic ASCII text, not a replacement for real-producer,
formatting-heavy, picture/object, compressed, legacy-code-page, malformed,
protected, external-link, cold-source, conversion, or broad edit/patch
matrices. Exact source save is intentionally one whole-source write today; the
benchmark records that behavior rather than claiming bounded chunking.

The next source-audited non-iWork candidates remain separate:

1. Measure ODT transaction snapshot handoff from the package's existing shared
   immutable bytes instead of cloning the complete ZIP.
2. Add six public semantic cases apiece for DOC, XLS and PPT before any further
   CFB ownership experiment: open, list, one object, full scan, no-op edit/save,
   and one edit/save.
3. Continue source-backed OPC query/edit/patch coverage and cache-budget work.

iWork remains deferred while the `iwa-*` crates are modified independently.
