# 0913 owned-source direct-read source review

Status: **conditional pass** pending the execution gates and the test/protocol
corrections listed below. This review is a static source and architecture
review; it did not build the workspace, run the ZIP tests, or run the PPTX
workload.

The reviewed candidate is the bounded change from parent
`66b944b07fcf4c6fa37f10e7ed83a02417455f57` in:

* `crates/soapberry-zip/src/office.rs`
* `crates/litchi-opc/src/payload.rs`

The reviewed seam is the one identified by the 0912 source audit: an
`IndexedArchive<Arc<Vec<u8>>>` may read an owned retained source directly,
while generic `ReaderAt` sources keep the existing `MemberSpan` optimization.
The applicable constraints are the accepted ADRs 0001, 0002, 0003, 0005,
0006, 0008, 0010, 0011, 0024, 0030 and 0031, together with the 0912 audit.

## Source findings

The production diff keeps the candidate narrow and preserves the intended
ownership boundary. `DeferredPayload::decode` is the only caller of
`IndexedArchive::read_owned`; it still obtains the archive through the source
`OnceLock`, performs the existing per-part limit check, charges the shared
aggregate actual-byte counter, and publishes the decoded vector through the
existing per-payload `OnceLock`. No source bytes, payload provenance, writer,
publication, or ordinary format API changes.

The specialized reader uses a private `DirectOwned` source policy and a
passthrough `SpannedSource` backed by the immutable `Arc<Vec<u8>>`. The generic
`IndexedReadSession::new` still selects `MemberSpan`, so file, range,
cancellable, versioned, and other `ReaderAt` sources retain their existing
read grammar. The direct method does not use `unsafe`, `set_len`, or an
uninitialized output buffer, and it does not introduce a cache or a shared
mutable decoder.

The direct path reuses the same `get_entry_from` local-header parser, local
ZIP64 resolution, source-boundary checks, output reservation, overrun sentinel,
Store and Deflate dispatch, `ZipVerifier`, data-descriptor parsing, accounting,
and final decoded-size check as the generic path. Since `Arc<Vec<u8>>` cannot
produce an external transport or cancellation error, removing its speculative
member-span read does not remove a caller-visible I/O observation. Malformed
local framing, truncated records, CRC failures, size failures, unsupported
methods, and source bounds still reach the same validation code.

The public `#[doc(hidden)]` method is low-level plumbing on the already
low-level archive type and is called only from `litchi-opc`; it does not leak
an archive type through the ordinary OPC or Office CRUD APIs. Its visibility is
needed for the crate boundary and is acceptable only with the existing
low-level API caveat: it is not a promise that callers should select physical
read policies themselves.

## Invariant review

* **Preservation and publication:** The candidate only changes how a deferred
  payload is decoded. It leaves the retained source archive, payload
  provenance, exact-source authorization, writer planning, raw member copy,
  and unchanged-member publication untouched. Existing lazy OPC tests and the
  new concurrency test source exercise exact no-op output and source sharing;
  these gates still need to execute in the isolated candidate source.
* **Limits:** Declared archive admission remains in the index constructor.
  `PartBytes` and `TotalPartBytes` are checked after verified decode exactly as
  before. The direct route does not reserve a second member-span allocation,
  but it retains the decoded output allocation and all existing checked
  reservations.
* **Concurrency and stable outcomes:** The source index and deferred payload
  outcome remain `OnceLock<Result<...>>`. Clones share the deferred payload
  cell and source handle. The added OPC concurrency cases cover concurrent
  first success, concurrent first corruption, stable refusal replication, one
  inflated-part charge, and exact source publication. Their execution is a
  required gate, not an inference from the type shape.
* **Generic-source behavior:** The default session constructor selects
  `MemberSpan`; no file, remote/range, cancellation, or versioned source is
  routed through `DirectOwned`. The candidate therefore does not broaden the
  immutable-memory assumption.
* **Architecture:** The change remains below the OPC owner in the
  `soapberry-zip` physical owner and does not add a facade or peer-format
  dependency. It preserves the downward dependency direction and the
  source-preserving lazy OPC design of ADR 0030.

There is a pre-existing ADR wording mismatch that the coordinator must record
before calling the clone gate complete. The implementation has
`PartPayload::Deferred(Arc<DeferredPayload>)`, so cloning a package shares the
`DeferredPayload` and its `OnceLock`; the new concurrency test correctly
asserts one decoded `Arc` and one aggregate charge across clones. ADR 0030's
accepted prose instead says that cloning a package clones each part's cell and
that independently decoded clones can be charged twice. This candidate does
not introduce that mismatch. The owner must amend the ADR or explicitly bind
the current shared-cell behavior as the governing contract; otherwise the
source and test evidence cannot be described as satisfying both statements.

## Verification conditions

Before accepting the candidate, the coordinator must retain passing evidence
for the following conditions:

1. Store and Deflate payloads, data descriptors, truncation, CRC, source
   bounds, limits, and ZIP64 **local-header** size metadata produce byte-equal
   results and equivalent refusal identities through the generic and direct
   routes.
2. Error comparisons check the complete relevant `ErrorKind` fields and
   stable diagnostic identity. The current helper no longer applies the
   blanket `is_content_fault()` assertion that failed for the v7 I/O wrapper,
   but it still compares the family plus `Debug` and `Display` only. The gate
   should retain typed field assertions for each refusal class rather than
   treating the family label or a formatted diagnostic as the contract.
3. Concurrent first forces share one successful decoded allocation and one
   aggregate charge; a failed first force is recorded and repeated unchanged.
4. An untouched owned package and a package with a forced payload publish the
   same bytes; source allocation identity and unchanged raw members remain
   intact.
5. The matched native baseline/candidate workload passes its semantic, source,
   output, sink, and untouched-member guards before any timing result is
   admitted.

The checked-in `zip64.zip` fixture exercises ZIP64 archive/central metadata,
but its local header carries ordinary 32-bit member sizes. The ordinary
differential test also has a descriptor-success case and CRC/truncated-Deflate
refusals, but it does not independently mutate a declared size or descriptor
record. The supplemental `owned_direct_adversarial.rs` fixture now supplies
forced local ZIP64 sizes, a 64-bit descriptor, declared-size mutations,
descriptor corruption, and local/source-bound failures; it must run through
the fixture admission step and restore the frozen source inventory before the
quality gate is considered complete.

No production source issue was found in the bounded design. The verification
conditions above are evidence requirements, not permission to weaken
validation or error matching. No latency, allocation, or end-to-end speedup
claim follows from this static review.
