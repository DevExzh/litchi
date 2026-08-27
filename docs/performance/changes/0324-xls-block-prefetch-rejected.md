# Change 0324: Reject XLS worksheet block prefetch before implementation

Status: rejected before implementation.

## Existing correctness boundary

`crates/litchi-xls/src/workbook/source.rs` currently scans a worksheet one
BIFF frame at a time. `WorksheetScan::next_frame` reads exactly the four-byte
header, validates the record length, worksheet boundary, record-count limit,
and scanned-byte limit, and only then advances the frame accounting.
Recognized records use `read_payload` for an exact payload read; unrecognized
records use `skip_payload` for an exact payload skip. Both paths check the
optional `ExecutionContext`, and the header path checks it before every frame.

The same source file deliberately keeps workbook-global handling separate:
the first pass reads only global headers and skips each payload, then the
semantic pass reads one exact bounded global range. Worksheet payloads are not
materialized during open. `crates/litchi-xls/tests/source_backed.rs` locks this
down with `raw_handoff_reads_global_headers_then_one_exact_global_range`,
`truncated_global_header_is_rejected_without_overread`, and the worksheet
limit and range assertions.

## Rejection rationale

A reusable block buffer would cross the current header/payload/skip
transaction boundary. It could read bytes belonging to a later frame before
the current frame's length, boundary, record-count, or byte-limit checks have
decided that the later bytes are admissible. It would also need to preserve
the distinction between payload reads and skips, exact overread rejection,
and the existing allocation bounds.

The owner freshness fence is independently observable. Metadata operations and
selected-cell queries call `ensure_current` around their work, while
`ensure_current_parts` validates the retained source around CFB access. A
prefetch transaction must not allow a source change to be hidden by bytes
already buffered. It must also preserve deferred worksheet errors: opening
validates workbook globals and worksheet offsets, but worksheet BOF version and
substream validation remains deferred until selected access, as covered by
`worksheet_bof_version_and_substream_type_are_deferred_to_selected_access`.

Finally, execution-aware scans check cancellation before frame headers, before
payload reads or skips, and during semantic cell processing. Moving multiple
frames into one read would change cancellation checkpoints and the point at
which read, limit, or malformed-frame errors are reported. No safe block size
or transaction protocol was established from source inspection, so no
production edit or timing claim is made.

## Safer next hypothesis

Investigate an exact-range CFB frame transaction: first retain the existing
per-frame validation and deferred-error decisions, then issue one explicitly
bounded CFB range read only for an already-admitted contiguous set of frames.
The design must retain source freshness fences, read limits, cancellation
checkpoints, exact EOF and boundary behavior, and the current payload-versus-
skip semantics before any implementation or benchmark is attempted.

