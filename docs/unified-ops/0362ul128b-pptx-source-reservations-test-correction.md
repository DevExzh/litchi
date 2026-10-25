# 0362ul128b — PPTX source reservation regression correction

Date: 2026-10-25

This follow-up corrects the source-only regression added by 128a. Chunk
admissions are now obtained through the real
`admit_shape_text_chunks`/`admit_shape_property_chunks` producers rather than
constructing the private `ChunkAdmission` fields in a test. The tight-budget
case now uses a 1 KiB parent budget and a 4 MiB child budget, proving the
hierarchical `ResourceLimit` refusal while the child-local memory ceiling is
large enough for the reservation.

It also adds a real source-backed slide-property edit with a valid output
ceiling only 128 bytes above the input and a managed transient envelope larger
than that ceiling. This verifies the consumer path, not only private
reservation arithmetic. Record 0152's output-before-memory guarantee remains
the text-box-specific contract; no adjacent candidate pre-staging was added.

No Cargo, build, rustdoc, metadata, or runtime command was run. The source
tree remains protected by the existing 128a production fix.
