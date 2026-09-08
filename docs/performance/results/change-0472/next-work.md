# Next bounded-streaming evidence gap

Independent read-only audit identifies a dedicated `docx_streaming_create`
harness selector as the next useful batch. Public `StreamingDocumentWriter`
already exists (`crates/litchi-docx/src/streaming.rs`) and has reopen, short-write,
invalid-input and cancellation tests. Its semantic scratch accounting does
not prove process RSS or total allocator bounds. Existing DOCX small-creation
selectors build an owned Package and output Vec, so they do not measure this
writer. Current XLSX/PPT measurements cannot fill that gap.

Measure fresh DOCX creation at 64 / 8,192 / 131,072 paragraphs with an explicit
64-byte scratch limit and the existing zero-retaining HashingDiscardSink.
Use untimed materialized preflight for reopen, topology, paragraph/text digest
and deterministic output checks. Measure operation allocator peak only around
the writer; report caller input, ZIP/sink state and process RSS separately.
No speculative production change is required before this evidence exists.

Keep the four append meanings separate: fresh creation, logical append to an
existing structure, adding a package Part, and arbitrary edits followed by
repackaging. This proposed DOCX case proves only fresh creation. PPTX needs
its own follow-up because StreamingPresentationWriter retains ZIP directory
and part-name metadata proportional to the number of parts. Existing XLSX
streaming evidence (change 0432) has a constant 420,110-byte incremental peak
across its three tested sizes, with scope limits. The broader goal stays open.
