# Change 0299: selected OPC part streaming

- **Status:** accepted
- **Date:** 2026-08-27
- **Scope:** deterministic bounded API/resource behavior
- **Performance claim:** none

## Contract

`litchi-opc::PartView` exposes `stream_to` and
`stream_to_with_accounting`. Both methods stream one already-admitted OPC part
into a caller-owned `Write` sink and return the exact number of decoded bytes
accepted by that sink.

The implementation routes the selected catalog entry through
`IndexedArchive::read_entry_to_with_accounting`. It does not call `PartView::data`,
enter or join the part cache, create `PartData`, or retain payload bytes in the
cache. A warm cache is intentionally ignored. A later `data()` call remains a
cold materialization, and repeated streaming repeats the physical archive work.

The already-admitted catalog `PartBytes` and `TotalPartBytes` policy remains the
admission gate. At stream time, execution and source-version state are checked
before metadata and payload I/O, the declared per-part limit is checked before
payload I/O, and aggregate total-part policy is not charged again. Managed
execution charges `Work` once by declaration, source reads use the existing
`InputBytes` accounting, and `OutputBytes` counts only bytes actually accepted by
the caller sink. No `Memory` or `Parts` reservation is created by streaming.

Cancellation and source-version checks run before, during, and after the
operation. `SourceChanged` takes precedence over competing lower-level errors.
The sink is neither flushed nor rolled back. A failure after accepted bytes uses
the existing `OpcError::IncompleteOutput { written, source }` representation;
zero-prefix failures retain their underlying typed error. A late CRC or source
failure may therefore report `written` equal to the declared size.

ZIP accounting is merged into the caller-owned `OpcOperationAccounting` on both
success and failure. Stored and deflated counters retain their compression-specific
meaning; raw or generated publication counters remain zero. Operation, source,
sink, and execution errors take precedence over accounting overflow. If merging
is the only failure, the accounting overflow is returned, wrapped as incomplete
output when bytes were accepted.

The selected-part sink records only local accepted bytes while it is running;
the OPC output counter is merged once after the stream and final source/context
fences. Accounting overflow therefore cannot manufacture a sink I/O failure or
erase the accepted-prefix count.

The low-level `IndexedArchive::read_entry_to_with_accounting` Store and Deflate
paths retain the observed compressed/stored source-byte count when streaming
fails. An accounting overflow on that error path does not replace the primary
stream error.

Execution and cancellation failures crossing the ZIP reader's I/O boundary
carry a private typed error source and are restored by the central OPC error
mapping layer. The source snapshot has no shared consumable execution-failure
slot, so concurrent operations cannot steal one another's typed failure.

## Evidence boundary

This record makes no claim about latency, throughput, allocation rate, total RSS,
OS or physical I/O, decoder internals, caller sink behavior, package-index or
catalog construction, concurrency, cache-wide behavior, or cross-format
serialization. The bounded claim is limited to the ZIP reader's fixed staging
buffer while streaming one selected, already-admitted part.

Focused tests cover stored/deflated accounting parity, accepted-prefix failures,
cache non-materialization, repeated reads, the typed execution-I/O round-trip,
and deferred output-accounting overflow. This pass adds no new claim for
declared-limit, cancellation, or source-mutation streaming cases. No benchmark
selector is added by this change.
