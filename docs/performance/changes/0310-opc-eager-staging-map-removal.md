# Change 0310: OPC eager staging-map removal

Status: implemented; focused validation complete

`performance_claim: none`

## Staging ownership

The eager serialization path no longer builds a name-keyed decompressed
`HashMap` with separately owned part-name keys. Decompressed payload ownership
is moved in serialization order as `Arc` values into the final ordered
serialized-part collection, preserving the package's existing part order and
ownership boundary without a second name-keyed staging index.

## Unchanged eager boundary

`read_many` continues to read all requested payloads into its all-payload
`Vec`, and the eager final package retains the same residency model. The
`OpenSession` lifecycle and its existing eager materialization boundary are
unchanged; this change does not turn the eager path into a streaming or
source-backed publication path.

## Validation and error behavior

Existing preflight checks remain before the staging transfer. Typed allocation,
read, size, and serialization errors retain their established precedence, and
relationship/content validation is not reordered or weakened by the staging
representation change. No new relationship interpretation or publication
semantics are introduced.

## Regression scope

Validation covered 244 OPC library tests, including 2 new focused staging-path
tests. Strict OPC library Clippy and rustdoc checks also passed. Final
repository handoff still expects rustfmt and diff-hygiene checks. This record
does not claim a complete package-shape matrix, every relationship family, or
publication coverage beyond the focused staging-path checks.

## Measurement boundary

No total-memory, RSS, peak-memory, or OOM improvement is claimed. The change
also does not claim a throughput improvement and does not alter the ordinary
eager final-residency or `OpenSession` contracts.
