# Change 0311: OPC fallible map growth

Status: implemented; focused validation complete

`performance_claim: none`

## Fallible allocation boundaries

The eager OPC path now uses fallible reservation before growing the eager part
map. Relationship unmarshalling and public relationship insertion likewise
reserve capacity fallibly, and the part-publication path uses `try_add_part`
instead of an infallible growth operation. Duplicate relationship-ID
preflight occurs before relationship-storage reservation, preserving the
existing invalid-input precedence without allocating for a duplicate.

These reservations convert allocation failure at the touched growth points into
the existing typed error path. They do not change the ownership of parts,
relationships, or their serialized order.

## Preserved behavior

Semantic decoding, validation and error precedence, relationship handling, and
transaction atomicity remain unchanged. A failed reservation or insertion
still leaves the candidate package unpublished; no partial eager package is
exposed through the existing API.

Compaction or redesign of retained provenance metadata is explicitly deferred
to a separate change. This record covers fallible growth only.

## Regression scope

Validation covered 244 OPC library tests. Strict OPC library Clippy and
rustdoc checks also passed. Final repository handoff still expects rustfmt and
diff-hygiene checks. This does not claim a complete allocation-failure matrix
for every eager map, relationship, or part-publication path.

## Measurement boundary

No total-memory, RSS, peak-memory, or OOM improvement is claimed. This change
also makes no throughput claim and does not alter the eager package's final
residency or semantic behavior.
