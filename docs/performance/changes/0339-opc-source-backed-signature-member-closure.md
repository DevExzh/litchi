# Change 0339: OPC source-backed signature member closure

## Scope

This change closes the source-backed OPC seam around package signatures. The
package reader performs a conservative raw-member signature scan before an
edit is published. Detection is based on authored package members and their
relationship/content-type metadata, including signature-origin and signature
members, rather than on a best-effort interpretation of the signature XML or
on a cryptographic verification result.

The scan treats the signature member closure as opaque package state. It does
not assume that a source-backed part can be edited safely merely because the
part itself was not materialized, because a signature may cover that part or
other package members reachable through the authored signature relationships.
Ambiguous, malformed, or unrecognised signature-shaped members are handled
conservatively. A false positive is preferable to publishing an edit whose
signature coverage is unknown.

The closure check is bounded by the existing package read, relationship,
member, and output limits. It does not expand arbitrary payloads, evaluate
signature transforms, or turn raw signature bytes into a public signing API.
Signature bytes, relationship ordering, and unrelated source-backed members
remain losslessly retained when no effective mutation occurs.

## Lossless and mutation boundaries

- An exact semantic no-op returns the original package bytes and an identity
  patch. It does not rewrite signature members, relationships, content types,
  source-backed parts, or package topology, even when the raw scan finds
  signature-shaped members.
- An effective edit crossing a package with a detected or ambiguous signature
  member closure is refused before any part, relationship, or manifest bytes
  are changed. The refusal is typed and atomic; it is not silently converted
  into a signature-stripping or unsigned rewrite.
- Materializing a source-backed package into an owning `OpcPackage` retains
  the conservative signature policy through opaque non-Part member metadata.
  Ordinary owned mutation therefore cannot bypass the source-backed guard.
- A nonempty topology plan cannot introduce a root signature Part, signature
  content type, signature relationship type, or internal signature target on
  an otherwise unsigned source. Explicit eager signing APIs remain the owner
  of signature creation.
- Ambiguous exact or ASCII-case-equivalent physical names are never resolved
  by selecting the first central-directory entry during changed publication.
  Exact source copy remains separate from this mutation refusal.
- A source-backed package without detected signature state may use the normal
  validated part-edit path. The edit still preserves all untouched raw
  members and relationships and applies the ordinary source freshness,
  cancellation, read, and output limits.
- Signature detection does not claim that a package is cryptographically
  valid, identify the signer, or prove the exact set of signed bytes. It only
  establishes whether mutation safety can be known from the bounded raw OPC
  member closure.
- Patches are source-preconditioned and reversible. A no-op patch is an
  identity operation. A refused signed-package mutation produces no forward
  patch and no partially rewritten package; inverse restoration is therefore
  not offered for that refusal.
- Signature origin parts, signature parts, signature relationships, and
  signature-related content-type entries remain opaque and preserved. They
  are not edited as ordinary application parts.

This change is limited to mutation gating and lossless raw-member handling.
It does not add signing, signature verification, signature removal, or
re-signing. It does not promise that a caller can modify a signed package by
editing the signature closure itself.

## Validation status

- `cargo fmt --package litchi-opc`: passed.
- The complete `source_backed_topology` integration target passed: 37 tests,
  zero failures, one test thread.
- The exact focused command was:

  ```sh
  CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0339-target \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
  cargo test -p litchi-opc --test source_backed_topology -- \
  --test-threads=1
  ```

- Independent read-only review closed the raw-member, owned-materialization,
  signature-introduction, physical-name ambiguity, and exact-no-op paths with
  no remaining P0/P1 finding.
- Broader OPC library/integration, strict Clippy, rustdoc, benchmark, and
  repository-wide gates were not run after the prior host OOM. No broader
  validation or performance claim is made.
- All Cargo work was serialized on one root-disk-backed target. No parallel
  rebuild and no `/tmp` or `/dev/shm` Rust target was used.
- Scoped diff checks passed and the isolated target was deleted. At
  finalization the root disk had 136 GiB free and `/dev/shm` used 53 MiB;
  the unrelated `/tmp` tmpfs was not modified.

## Performance claims

`performance_claim: none`

No latency, throughput, allocation, RSS, decompression, or process-memory
claim is made. Bounded raw-member scanning, exact no-op publication, and
mutation refusal are correctness and preservation properties, not measured
performance results.

## Follow-up

Future signing work must define the supported signature profiles, canonical
member closure, transform and digest semantics, key handling, and explicit
strip/resign policy. Signature removal must be an opt-in, auditable operation
with a documented effect on origin parts, relationships, content types, and
all covered members. Re-signing must rebuild and validate the complete closure
atomically; neither behavior is implied by this source-backed mutation guard.
