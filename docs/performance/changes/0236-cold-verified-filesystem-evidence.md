# Change 0236: strict Linux cold-verified filesystem evidence

Date: 2026-08-20

Status: harness capability; no measurement or performance claim

## Scope

The standalone `tools/perf-baseline` harness now accepts the opt-in
`cold-verified` cache state.  The existing `warm` and `cold-requested` states,
including their default selection and advisory semantics, are unchanged.

The verifier is deliberately Linux-only and fails closed.  It requires a
regular, non-empty, page-aligned source opened read-write.  Filesystem
admission uses `fstatfs` on that source FD and a numeric magic allowlist:
`0xef53` for ext2/3/4, `0x58465342` for XFS, `0x9123683e` for Btrfs,
`0xf2f52010` for F2FS, and `0x2fc12fc1` for ZFS.  It `fsync`s the source,
requests `posix_fadvise(DONTNEED)`, and invokes the canonical, hashed,
versioned external util-linux `fincore` command with an exact JSON column
list.  One strict record must prove zero resident, dirty, and writeback bytes
immediately before the timed source-touching operation.  Raw stderr and
method/fallback evidence are retained; unrecognized fallback is rejected.
The child then requires a positive `/proc/self/io` `read_bytes` delta.
Prepared query controls are excluded and report
`ineligible_prepared_query_control` rather than receiving a misleading timed
result.

The harness uses a private page-aligned source copy.  ZIP alignment extends
the EOCD comment field so logical members remain unchanged; CFB alignment uses
trailing zeroes after the declared sector chain.  Each proof records the
aligned source SHA-256 and size, numeric filesystem magic, and canonical
fincore identity.  DOCX verification compares the aligned archive's derived
semantic signature with the unpadded corpus signature while retaining the
aligned physical size as a separate identity fact.  Source paths and device
identifiers are not serialized.

## Status and claim boundary

Every requested filesystem case records `cold_verified_status`.  Ineligible
hosts or proofs (for example unsupported filesystem, unavailable or malformed
`fincore`, resident/dirty/writeback pages, and zero process `read_bytes`) do
not produce a `cold-verified` `CaseResult`.  The evidence claim is limited to
page-cache residency/dirty/writeback proof and process `read_bytes`; it does
not establish physical-media temperature, device-cache state, or storage
latency.  No unsafe code or production dependency is introduced.

The feature depends on the host providing Linux procfs, `getconf`,
`posix_fadvise`, and util-linux `fincore`, plus a supported block-backed
filesystem.  These are external host prerequisites and are not verified by
the repository-only checks.
