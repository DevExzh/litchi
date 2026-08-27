# Change 0334: OPC source-backed physical-Part deletion

## Scope

This change adds bounded, low-level source-backed operations for physically
deleting OPC Parts, optionally alongside payload replacements, without
materializing deleted payloads. It is a physical packaging primitive: it omits
the selected ZIP members and preserves unrelated members whenever the source
layout permits a lossless ZIP32 rewrite. It does not infer or repair package
semantics.

The caller must update `[Content_Types].xml`, remove inbound relationships, and
remove any owned `.rels` member that becomes obsolete. A caller that omits a
Part without updating those owners can produce an invalid OPC graph; this
primitive never guesses those edits.

## Implementation

- `write_part_overlays_with_deletions_to_stream` and its shared-payload
  counterpart accept at most 64 combined replacements and deletions.
- Every selected `PackURI` resolves through the immutable source catalog before
  output. Missing, duplicate, case-equivalent duplicate, and replace/delete
  overlap selections are typed pre-output errors.
- Deletions map the exact physical member name to
  `soapberry_zip::PreservationAction::Omit`; deleted payload bytes are not read
  or decompressed.
- An empty plan copies the complete source artifact byte for byte. A changed
  signed source is refused before output.
- Untouched local records, compressed payloads, extras, comments, ordering, and
  central records are preserved. Central local-offset fields are recalculated
  when omission shifts later members.
- The preservation index now uses the package's configured ZIP limits rather
  than the default profile.
- Source-version, cancellation, execution-context, output-budget, and partial
  sink semantics are inherited from the existing bounded sequential publisher.
- Unsupported ZIP64, trailing-byte, multi-disk, prefixed, overlapping, and
  ambiguous layouts remain typed refusals; there is no eager fallback.
- The existing replacement-only APIs retain their behavior and delegate with an
  empty deletion set.

## Caller contract

The operation owns only physical ZIP member omission. A semantic caller must
compose all required graph changes in the same consuming publication plan:

- update or remove the relevant `[Content_Types].xml` mapping;
- remove every inbound relationship to the deleted Part;
- remove an owned `.rels` member when appropriate; and
- update format-specific indexes and references.

The API is a consuming low-level publisher, not a graph-safe format deletion or
a standalone reversible patch object.

## Validation evidence

All Cargo commands used `CARGO_BUILD_JOBS=1` and the isolated
`/dev/shm/litchi-0334-target` directory.

- Focused `source_backed_topology` integration suite: 26 passed.
- Complete `litchi-opc` library suite: 264 passed.
- Complete `litchi-opc` integration suites: 67 passed.
- Strict `cargo clippy -p litchi-opc --lib --tests -- -D warnings`: passed.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-opc --no-deps`: passed.
- New tests cover empty-plan byte identity, single deletion, combined
  replacement/deletion publication, untouched raw-member preservation,
  unchanged content-type and relationship bytes, typed selection refusals
  before sink output, and signed-source refusal.
- Existing shared publication-path tests continue to cover source-version
  changes, cancellation, partial sink failures, configured resource limits,
  ZIP32/layout refusals, and bounded sequential output.

## Performance claims

`performance_claim: none`

No benchmark, latency, throughput, RSS, allocation, or process-wide memory
claim is made. The architectural work removed mandatory payload reads from the
physical deletion mechanism, but no dedicated before/after deletion benchmark
was run.

## Follow-up

This is a prerequisite for source-backed XLSX calculation-chain cleanup. The
XLSX owner may omit `xl/calcChain.xml` only after transactionally updating the
workbook relationships and content types in the same publication, while
retaining the source, signature, limit, and preservation guarantees above.

## Residuals

- This primitive does not decide whether a Part is semantically safe to delete
  and does not discover format-specific references.
- Deleting a Part does not automatically delete its relationship member.
- The conservative output bound does not subtract omitted spans and can refuse
  a near-ZIP32-limit source even when deletion would make the result fit.
