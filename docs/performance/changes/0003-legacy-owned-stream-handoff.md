# Legacy XLS/PPT owned stream handoff

Status: partially accepted after measurement
Production base: `2665d572b78f0b3efd9ecfc4bd1fda09f8786ae3`

## Mechanism and corpus

Fresh XLS and PPT writers now move their already-owned generated stream
buffers into `OleWriter::create_stream_owned`. Stream/storage insertion order,
CFB bytes, encryption timing, and public APIs are unchanged.

The deterministic `payload-heavy` corpus exercises real public writer APIs:

- DOC: 128 paragraphs producing a 5.13 MB `WordDocument` stream;
- XLS: 128 worksheets with one legal 32,700-byte string each, producing a
  4.22 MB `Workbook` stream;
- PPT: 16 slides and 128 40,000-byte text boxes, producing a 5.15 MB
  `PowerPoint Document` stream.

Matched release binaries differed only in the DOC/XLS/PPT handoff. Each ABBA
replicate used 150 samples after 10 warm-ups. All before/after CFB archive
SHA-256 hashes are identical. Raw reports are the four
`results/abba-legacy-owned-heavy-matched-*.json` files.

| Writer | Before p50 | After p50 | p50 change | Mean change | Decision |
|---|---:|---:|---:|---:|---|
| DOC | 4.357 ms | 6.902 ms | +58.42% | +58.87% | Rejected and reverted |
| XLS | 4.126 ms | 4.065 ms | -1.48% | -0.15% | Retained for peak-memory reduction |
| PPT | 6.312 ms | 5.035 ms | -20.23% | -19.96% | Retained |

Heaptrack used 50 payload-heavy iterations:

| Writer | Peak heap before | Peak heap after | Profiler RSS before | Profiler RSS after |
|---|---:|---:|---:|---:|
| DOC | 36.82 MB | 38.77 MB | 39.12 MB | 36.07 MB |
| XLS | 44.73 MB | 40.50 MB (-9.5%) | 47.80 MB | 41.77 MB (-12.6%) |
| PPT | 41.46 MB | 36.32 MB (-12.4%) | 40.34 MB | 35.12 MB (-12.9%) |

The DOC builder grows a buffer with substantial spare capacity. Moving that
allocation keeps the oversized capacity live throughout CFB serialization;
the old exact-sized copy is materially faster for this corpus. In accordance
with the regression gate, the DOC production change and its change-specific
test were removed. This is evidence that ownership transfer is not
automatically a win when producer capacity is significantly above length.

XLS latency is neutral, but removing the 4.22 MB copy reduces peak tracked heap
by 4.23 MB and profiled RSS by 6.03 MB. PPT improves both latency and memory.
Both retained formats pass their all-feature writer suites and byte-exact
determinism checks.
