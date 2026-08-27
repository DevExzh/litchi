# Change 0300: XLS source-backed lazy SST resolution

status: implemented
performance_claim: none

## Scope

The source-backed XLS owner retains a bounded immutable shared-string locator:
SST/CONTINUE segment source ranges and fallibly allocated per-entry spans. It
does not retain a decoded `Vec<String>`, a complete raw SST copy, a decoded
string cache, or rich-string properties. A selected `LabelSst` cell opens an
independent `SharedOleStreamCursor`, reads only the entry span, and decodes an
owned value for that result.

## Correctness contract

Open still scans every SST entry with the BIFF8 continuation state machine. It
preserves count validation, `cstTotal >= cstUnique`, source SST limits, UTF-16
code-unit boundaries, continuation width flags, rich-text runs, `ExtRst`,
truncation, and the existing empty/trailing `CONTINUE` compatibility. Selected
lookups validate the `u32` index before any SST source read and preserve the
existing invalid-index and unavailable-SST `CellValue::Error` text.

Each lookup uses a fresh cursor and shares only immutable locator metadata.
Execution checks and source freshness fences remain active around source reads,
and the final query fence preserves `SourceChanged` precedence. The eager
`Workbook` SST and rich-property APIs are unchanged.

## Allocation and evidence boundary

This is a bounded retained-allocation-shape improvement, not a measured
performance claim. Open may still scan and decode each SST value transiently,
and the owner retains the CFB handle, workbook catalog, global metadata, and
locator. The focused evidence is retained-table absence and selected replay
parity; the split fixture exercises an SST-to-`CONTINUE` transition, not
multi-entry range locality. This note excludes total RSS, peak allocator
behavior, I/O volume, latency, throughput, eager parsing, editor behavior, and
public rich-property coverage.
