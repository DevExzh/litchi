# 0912 owning cold deferred-payload source audit

This is a source-only audit of the owning cold path used by the retained 0910
snapshot binary. It traces the first `DeferredPayload::decode` through the
indexed ZIP reader, identifies the extra source-span work that an immutable
owned source does not need, and records one bounded candidate. No production
code, binary, protocol, workload, or profile was changed or run for this
audit.

## Frozen custody

The source is bound to
`docs/performance/results/change-0910/source-v1.json` (1,756,112 bytes,
SHA-256
`7a6d1310fb97dde1f275ec8702efabbef2cfb643c9ee4c1178169fed53aa4586`,
parent `468fe2931d6532c32f4c6608590816b8eaa88e31`). The retained executable
is `docs/performance/results/change-0910/binaries/native` (9,301,456 bytes,
SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`).
The current files below match the frozen source receipt:

| source file | SHA-256 |
| --- | --- |
| `crates/litchi-opc/src/payload.rs` | `fc6e468322dcc2939b5eeb2d88ca1401fdb5c93849a131947e3e36de85779998` |
| `crates/litchi-opc/src/phys_pkg.rs` | `3dfa9ec92c0d090b0bc753ecacc63f263975a470c267b79b03ae1c5898ac3ab8` |
| `crates/litchi-opc/src/pkgreader.rs` | `d3ae9f5057a851fc190ffcc8070f6c9ed7cbd221c15d262e68886f0c64e62040` |
| `crates/soapberry-zip/src/office.rs` | `f72bfef6e6a072b3e88c465afb1358e4f77046e37c606ed3902fe3a1621ceac3` |
| `crates/soapberry-zip/src/archive.rs` | `f3496a11c6e7ed798f11396fce9a4a336f0a9d9ff167fc94b9bded14cb547fa5` |
| `crates/soapberry-zip/src/reader_at.rs` | `7d58b13c16007d830b8eac6ecc7e756918ec17e3e841d92c5b6b592157b9b482` |

The behavioural constraints are those in ADR 0030: the declared part-byte
preflight remains at open, actual `PartBytes` and `TotalPartBytes` checks stay
at first decode, a failure is recorded in its `OnceLock`, and retained source
provenance remains authoritative. The 0911 source audit and candidate
assessment are historical evidence only; this document does not reinterpret
their guest attribution as native timing.

## Call chain

The owning cold route is:

```text
owned Arc<Vec<u8>>
  -> DeferredPartSource::index (OnceLock<IndexedArchive<Arc<Vec<u8>>>>)
  -> DeferredPayload::force / decode
  -> IndexedArchive::read (fresh IndexedReadSession)
  -> member_source (optional zero-filled MemberSpan)
  -> get_entry_from (local framing and ZIP64 metadata)
  -> ZipReader + Store or Deflate decoder
  -> ZipVerifier (size, CRC, data descriptor)
  -> output Vec<u8>
  -> actual part and aggregate limits
  -> Arc<Vec<u8>> moved into the payload OnceLock
```

`DeferredPartSource` retains one `Arc<Vec<u8>>` and one shared index cell
(`crates/litchi-opc/src/payload.rs:37-114`). The index is built only when a
payload is first forced. `DeferredPayload::decode`
(`payload.rs:178-193`) calls `archive.read(&member)`, checks the actual output
length against `PartBytes`, atomically charges `TotalPartBytes`, and then
wraps the returned vector with `Arc::new`. That `Arc::new` adopts the returned
`Vec`; it does not copy the decoded payload. `force` records either that Arc or
the first error in a per-part `OnceLock`, so repeated access does not repeat
inflation or change error identity.

The declared-size pass remains earlier in deferred admission
(`crates/litchi-opc/src/pkgreader.rs:1495-1560, 1702-1740`). Structural
content types, relationships, encryption admission, and part classification
use the physical reader before ordinary payloads are deferred
(`pkgreader.rs:950-1026`).

There is one separate index-build repetition. Physical ingress constructs a
`LazyArchiveReader` whose `ArchiveReader::new_with_limits` parses and indexes
the central directory (`crates/litchi-opc/src/phys_pkg.rs:295-307`,
`crates/soapberry-zip/src/office.rs:2851-2965`). The first deferred decode
then constructs a new `IndexedArchive<Arc<Vec<u8>>>` from the same retained
bytes (`payload.rs:99-114`, `office.rs:3969-3984`). The second index is needed
by the current owned, positionally readable lifetime and entry-ID type, but
its central-directory scan and metadata validation are a real repeated setup
cost. Sharing or transferring that index would be a broader ownership/API
change; it is recorded as a follow-up risk rather than the bounded candidate
below.

## Where the extra copy occurs

`IndexedArchive::read` and `read_with_accounting` create a new
`IndexedReadSession` for each one-shot member read
(`crates/soapberry-zip/src/office.rs:4737-4756, 5030-5046`). The session has a
reusable Deflate decoder, but `DeferredPayload::decode` does not retain a
session between parts. Within the read, `member_source` computes a bounded
local-record window (`office.rs:5560-5615`), reserves a fresh `Vec`, fills the
whole allocation with zeroes, and reads the same bytes from its source into
that allocation (`office.rs:5633-5656`).

For the deferred source the reader is `Arc<Vec<u8>>`; its `ReaderAt`
implementation delegates to the underlying slice and copies directly into a
caller-provided buffer (`crates/soapberry-zip/src/reader_at.rs:285-297`). Thus,
when the window is admitted, the path performs this unnecessary work before
decoding:

1. allocate a temporary `MemberSpan` buffer;
2. zero-initialize its complete window with `resize(length, 0)`; and
3. copy the retained source bytes into that buffer.

`get_entry_from` then parses the local fixed header from the span and may
allocate and zero-fill a second small `local_extra` buffer for ZIP64 sentinel
fields (`crates/soapberry-zip/src/archive.rs:2090-2161`). If the span fully
covers the compressed payload, `SpannedSource::read_at` copies the compressed
bytes from the span into the ZIP reader's requested buffer
(`office.rs:1590-1604, 1662-1687`) before Store verification or Deflate
inflation. If the window is only partial, the payload correctly falls back to
the original source, but the temporary span allocation, zero-fill, and
prefetch copy have already happened.

The indexed-reader code does not explicitly zero-fill its output vector; it uses
`try_reserve_exact` and `read_to_end` (`office.rs:1818-1905`). Store members
copy decompressed bytes into the returned output, and Deflate members
copy compressed input through the current decoder interface. These remain
part of the current API path; this audit does not prove that all such copies
are fundamentally necessary or attribute initialization inside `read_to_end`. The
avoidable operation is the intermediate source-to-`MemberSpan` copy and its
zero-fill on an immutable in-memory source.

The read still has to run the existing framing and verification grammar.
`get_entry_from` checks local bounds and ZIP64 framing; the Store and Deflate
branches use `ZipVerifier`; and `ZipVerifier` checks decompressed size, CRC,
and data descriptors while avoiding only an identical second end observation
(`archive.rs:2422-2509`). The repeated local parse is intentional today: it
is the same parse and refusal order against either a buffered or passthrough
source. The ZIP64 `local_extra` zero-fill is a smaller, conditional instance
of the same pattern, not the main fixture-wide opportunity.

## Bounded candidate

Add an explicit internal owned-source read path for
`IndexedArchive<Arc<Vec<u8>>>` that uses the retained Arc directly as the
`ReaderAt` for `get_entry_from`, `ZipReader`, and `ZipVerifier`, bypassing
`member_source`'s speculative `MemberSpan`. The generic `ReaderAt` path keeps
the current span optimization, because file, cancellable, or versioned
sources can benefit from one bounded read and must retain their current read
and error grammar. The owned path has immutable in-memory bytes, so its
direct positional reads have no external I/O or cancellation side effect to
preserve.

The candidate is bounded if it keeps the existing indexed entry, local-header
parse, `read_limit` overrun sentinel, output reservation, Store/Deflate
decoder, `ZipVerifier`, and data-descriptor handling unchanged. It must also
retain all of the following:

- declared-size admission in `PackageReader`, then actual per-part and
  aggregate charges after the returned vector is verified;
- the shared source index `OnceLock` and each payload's success-or-error
  `OnceLock`, including stable error replication;
- exact retained source bytes and the existing ownership/provenance handle;
- source bounds and malformed-header/CRC/size refusal identities; and
- the returned decoded allocation, which later readers and publication use.

The safe first version should bypass the temporary span rather than introduce
an unsafe `set_len`/uninitialized-buffer reader. A future zero-fill-only
variant would need to prove initialization and truncation on every short or
failed read, and must preserve the current passthrough fallback. The owned
direct path should be compared against the current path for Store and Deflate
members, truncated records, ZIP64 local metadata, data descriptors, CRC and
limit refusals, concurrent first forces, source preservation, and byte-identical
publication before implementation is considered.

The existing reusable `IndexedReadSession` is a secondary opportunity:
`office.rs:1758-1795, 1905-1944` already reuses a Deflate decoder within a
session, while one-shot `archive.read` creates a fresh session per payload.
An operation-scoped batch session could reduce decoder setup on a serial
fingerprint, but it must not become shared mutable state across independent
`OnceLock` forces or alter first-error ordering. It is therefore deferred
until a measured native attribution and an API/concurrency proof exist.

No SHA backend conclusion follows from this source map. The native dispatch
question remains governed by the separate 0912 native evidence protocol.
